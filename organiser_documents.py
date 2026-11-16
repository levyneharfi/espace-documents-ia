"""À enregistrer à la racine du projet sous organiser_documents.py.
Arrêter le serveur puis exécuter : python organiser_documents.py
Ne modifie pas les clés, signatures ou réglages d'authentification.
"""
import ast
import json
import shutil
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent
DOCUMENTS = RACINE / 'documents_generes'
FORMATS = ('pdf', 'docx', 'xlsx', 'pptx')

HELPER = '''"""Résolution des documents classés, avec compatibilité des anciens chemins."""
from pathlib import Path

RACINE_DOCUMENTS = (Path(__file__).resolve().parent / 'documents_generes').resolve()
FORMATS = {'pdf', 'docx', 'xlsx', 'pptx'}


def dossier_format(extension):
    extension = extension.lstrip('.').lower()
    if extension not in FORMATS:
        raise ValueError('Format non pris en charge.')
    dossier = RACINE_DOCUMENTS / extension
    if dossier.is_symlink():
        raise ValueError('Dossier symbolique interdit.')
    dossier.mkdir(parents=True, exist_ok=True)
    return dossier


def trouver_document(fichier, extensions=None):
    if not isinstance(fichier, str) or not fichier or '\\\\' in fichier:
        raise ValueError('Chemin de document invalide.')
    relatif = fichier.removeprefix('documents_generes/')
    parties = relatif.split('/')
    if any(p in ('', '.', '..') for p in parties):
        raise ValueError('Chemin de document invalide.')
    if len(parties) not in (1, 2):
        raise ValueError('Chemin de document invalide.')
    nom = parties[-1]
    if nom.startswith('.'):
        raise ValueError('Nom de document invalide.')
    extension = Path(nom).suffix.lower()
    type_fichier = extension.lstrip('.')
    if type_fichier not in FORMATS:
        raise ValueError('Format non pris en charge.')
    if extensions is not None and extension not in extensions:
        raise ValueError('Format non autorisé pour cet outil.')
    if len(parties) == 2 and parties[0] != type_fichier:
        raise ValueError('Le dossier ne correspond pas au format.')
    candidats = [RACINE_DOCUMENTS / type_fichier / nom,
                 RACINE_DOCUMENTS / nom]
    valides = []
    for chemin in candidats:
        if chemin.parent.is_symlink() or chemin.is_symlink():
            raise ValueError('Lien symbolique interdit.')
        if chemin.is_file():
            if chemin.resolve().parent not in (
                RACINE_DOCUMENTS, RACINE_DOCUMENTS / type_fichier
            ):
                raise ValueError('Document hors du dossier autorisé.')
            valides.append(chemin)
    if len(valides) > 1:
        raise ValueError('Nom ambigu : plusieurs documents portent le même nom.')
    if not valides:
        raise ValueError('Document introuvable.')
    return valides[0]
'''


def remplacer(texte, ancien, nouveau, fichier):
    if texte.count(ancien) != 1:
        raise ValueError(
            f'{fichier} : passage absent ou ambigu, aucun changement appliqué.\n'
            f'{ancien!r}'
        )
    return texte.replace(ancien, nouveau, 1)


def remplacer_fonction(texte, nom, nouveau):
    arbre = ast.parse(texte)
    fonctions = [n for n in arbre.body
                 if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
                 and n.name == nom]
    if len(fonctions) != 1 or fonctions[0].decorator_list:
        raise ValueError(f'Fonction non décorée attendue : {nom}')
    fonction = fonctions[0]
    lignes = texte.splitlines(keepends=True)
    return (''.join(lignes[:fonction.lineno - 1]) + nouveau.rstrip() + '\n'
            + ''.join(lignes[fonction.end_lineno:]))


def preparer_code():
    modules = ['outils_pdf.py', 'outils_documents.py', 'outils_excel.py',
               'outils_pptx.py', 'outils_apercus.py', 'controle_pdf.py',
               'telechargements.py']
    codes = {}
    for nom in modules:
        fichier = RACINE / nom
        if not fichier.is_file() or fichier.is_symlink():
            raise ValueError(f'Fichier requis absent ou symbolique : {nom}')
        codes[nom] = fichier.read_text(encoding='utf-8')
    if (RACINE / 'rangement_documents.py').exists():
        raise ValueError('rangement_documents.py existe déjà : ne pas relancer la migration.')

    for nom, extension in [('outils_pdf.py', 'pdf'),
                           ('outils_documents.py', 'docx'),
                           ('outils_excel.py', 'xlsx'),
                           ('outils_pptx.py', 'pptx')]:
        texte = codes[nom]
        arbre = ast.parse(texte)
        affectations = [n for n in arbre.body if isinstance(n, ast.Assign)
                        and any(isinstance(t, ast.Name) and t.id == 'DOSSIER'
                                for t in n.targets)]
        if len(affectations) != 1:
            raise ValueError(f'{nom} : affectation DOSSIER inattendue.')
        n = affectations[0]
        lignes = texte.splitlines(keepends=True)
        texte = (''.join(lignes[:n.lineno - 1])
                 + 'from rangement_documents import dossier_format, trouver_document\n'
                 + f'DOSSIER = dossier_format("{extension}")\n'
                 + ''.join(lignes[n.end_lineno:]))
        # Seuls les chemins publics des livrables changent ; les noms restent identiques.
        texte = texte.replace('documents_generes/{pdf.name}',
                              'documents_generes/pdf/{pdf.name}')
        texte = texte.replace('documents_generes/{fichier.name}',
                              f'documents_generes/{extension}/{{fichier.name}}')
        texte = texte.replace('documents_generes/{f.name}',
                              f'documents_generes/{extension}/{{f.name}}')
        if extension == 'pdf':
            texte = remplacer(
                texte, 'DOSSIER_SOURCES = DOSSIER / "source"',
                'DOSSIER_SOURCES = DOSSIER.parent / "source" / "pdf"', nom)
            texte = texte.replace('DOSSIER_SOURCES.mkdir(exist_ok=True)',
                                  'DOSSIER_SOURCES.mkdir(parents=True, exist_ok=True)')
            texte = remplacer_fonction(
                texte, 'retrouver_pdf',
                'def retrouver_pdf(fichier: str) -> Path:\n'
                '    return trouver_document(fichier, {".pdf"})\n')
        elif extension in ('xlsx', 'pptx'):
            texte = remplacer(
                texte, "SOURCES = DOSSIER / 'source'",
                f"SOURCES = DOSSIER.parent / 'source' / '{extension}'", nom)
            texte = texte.replace('SOURCES.mkdir(exist_ok=True)',
                                  'SOURCES.mkdir(parents=True, exist_ok=True)')
            texte = texte.replace('documents_generes/source/{chemin_source.name}',
                                  f'documents_generes/source/{extension}/{{chemin_source.name}}')
            texte = texte.replace('documents_generes/source/{source.name}',
                                  f'documents_generes/source/{extension}/{{source.name}}')
        codes[nom] = texte

    nom = 'telechargements.py'
    texte = codes[nom]
    texte = remplacer_fonction(texte, '_fichier_autorise', '''def _fichier_autorise(nom: str):
    # Les URL continuent à signer le nom seul : anciens liens conservés.
    if not nom or '/' in nom or '\\\\' in nom or nom.startswith('.'):
        return None
    from rangement_documents import trouver_document
    try:
        return trouver_document(nom, set(TYPES))
    except ValueError:
        return None
''')
    codes[nom] = texte

    nom = 'outils_apercus.py'
    texte = codes[nom]
    texte = remplacer_fonction(texte, '_trouver', '''def _trouver(fichier: str) -> Path:
    from rangement_documents import trouver_document
    chemin = trouver_document(fichier, EXTENSIONS)
    if chemin.stat().st_size > 30_000_000:
        raise ValueError('Document limité à 30 Mo pour les aperçus.')
    return chemin
''')
    texte = remplacer(texte, '    dossier = CACHE / cle',
                      "    dossier = CACHE / origine.suffix.lstrip('.') / cle", nom)
    texte = remplacer(texte, '    dossier.mkdir(exist_ok=True)',
                      '    dossier.mkdir(parents=True, exist_ok=True)', nom)
    texte = remplacer(
        texte, "'fichier': 'documents_generes/' + origine.name,",
        "'fichier': origine.relative_to(RACINE).as_posix(),", nom)
    codes[nom] = texte

    nom = 'controle_pdf.py'
    codes[nom] = remplacer_fonction(codes[nom], '_trouver', '''def _trouver(fichier: str) -> Path:
    from rangement_documents import trouver_document
    return trouver_document(fichier, {'.pdf'})
''')
    codes['rangement_documents.py'] = HELPER
    for nom, texte in codes.items():
        compile(texte, nom, 'exec')
    return codes


def preparer_deplacements():
    mouvements = []
    ignores = []
    if DOCUMENTS.is_symlink():
        raise ValueError('documents_generes ne doit pas être symbolique.')
    for extension in FORMATS:
        for fichier in DOCUMENTS.glob(f'*.{extension}'):
            if fichier.is_symlink() or not fichier.is_file():
                raise ValueError(f'Fichier inattendu : {fichier}')
            mouvements.append((fichier, DOCUMENTS / extension / fichier.name))
    # Sources anciennes à la racine ou dans source/.
    for dossier in (DOCUMENTS, DOCUMENTS / 'source'):
        if dossier.is_symlink():
            raise ValueError(f'Dossier symbolique interdit : {dossier}')
        for fichier in dossier.glob('*.source.json'):
            if fichier.is_symlink() or not fichier.is_file():
                raise ValueError(f'Source inattendue : {fichier}')
            if fichier.stat().st_size > 20_000_000:
                raise ValueError(f'Source trop grosse pour cette migration : {fichier}')
            extension = next((e for e in FORMATS
                              if fichier.name.endswith(f'.{e}.source.json')), None)
            if extension is None:
                try:
                    contenu = json.loads(fichier.read_text(encoding='utf-8'))
                except (ValueError, UnicodeError):
                    ignores.append(str(fichier))
                    continue
                if isinstance(contenu, dict) and 'html' in contenu:
                    extension = 'pdf'
            if extension is None:
                ignores.append(str(fichier))
                continue
            cible = DOCUMENTS / 'source' / extension / fichier.name
            mouvements.append((fichier, cible))
    cibles = set()
    for origine, cible in mouvements:
        if cible.exists() or cible.is_symlink() or cible in cibles:
            raise ValueError(f'Conflit : {cible}. Aucun fichier ne sera écrasé.')
        cibles.add(cible)
        for parent in cible.parents:
            if parent == RACINE:
                break
            if parent.is_symlink():
                raise ValueError(f'Dossier symbolique interdit : {parent}')
    return mouvements, ignores


def main():
    codes = preparer_code()
    mouvements, ignores = preparer_deplacements()
    sauvegarde = RACINE / ('sauvegardes_rangement_' +
                           datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    sauvegarde.mkdir()
    for nom in codes:
        original = RACINE / nom
        if original.exists():
            shutil.copy2(original, sauvegarde / nom)
    manifeste = [{'avant': str(a.relative_to(RACINE)),
                  'apres': str(b.relative_to(RACINE))} for a, b in mouvements]
    (sauvegarde / 'deplacements.json').write_text(
        json.dumps(manifeste, ensure_ascii=False, indent=2), encoding='utf-8')
    deplaces = []
    ecrits = []
    try:
        for extension in FORMATS:
            (DOCUMENTS / extension).mkdir(parents=True, exist_ok=True)
            (DOCUMENTS / 'source' / extension).mkdir(parents=True, exist_ok=True)
            (DOCUMENTS / 'apercus' / extension).mkdir(parents=True, exist_ok=True)
        for origine, cible in mouvements:
            origine.rename(cible)
            deplaces.append((origine, cible))
        for nom, contenu in codes.items():
            (RACINE / nom).write_text(contenu, encoding='utf-8')
            ecrits.append(nom)
    except Exception:
        # Restauration de secours pour les erreurs Python ordinaires.
        # Une interruption brutale ou une panne disque peut nécessiter une reprise manuelle.
        for nom in reversed(ecrits):
            copie = sauvegarde / nom
            if copie.exists():
                shutil.copy2(copie, RACINE / nom)
            elif nom == 'rangement_documents.py':
                (RACINE / nom).unlink(missing_ok=True)
        for origine, cible in reversed(deplaces):
            if cible.exists() and not origine.exists():
                cible.rename(origine)
        raise
    print('Migration terminée ; syntaxe des fichiers modifiés vérifiée.')
    print(f'{len(deplaces)} fichier(s) déplacé(s), sans renommage des documents.')
    print('Sauvegarde du code et manifeste :', sauvegarde.name)
    print('Anciens chemins simples acceptés ; signatures de téléchargement inchangées.')
    print('Les anciens caches d’aperçus sont conservés ; les nouveaux seront classés par format.')
    if ignores:
        print('Sources non identifiées, laissées en place :')
        for nom in ignores:
            print(' -', nom)
    print('Redémarrer le serveur, puis tester création, téléchargement, source et aperçu.')
    print('Les modules Apple ne sont pas modifiés : ils n’ont pas été fournis.')


if __name__ == '__main__':
    try:
        main()
    except Exception as erreur:
        raise SystemExit(f'Migration interrompue : {erreur}')