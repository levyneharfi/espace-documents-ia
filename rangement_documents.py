"""Résolution des documents classés, avec compatibilité des anciens chemins."""
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
    if not isinstance(fichier, str) or not fichier or '\\' in fichier:
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
