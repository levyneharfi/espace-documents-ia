from images_documents import image_png
"""Création XLSX structurée, sans accès à des fichiers fournis par le modèle."""
import json
import math
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from time import perf_counter
from typing import Literal
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from pydantic import BaseModel, Field

from serveur_mcp import mcp

from rangement_documents import dossier_format, trouver_document
DOSSIER = dossier_format("xlsx")
SOURCES = DOSSIER.parent / 'source' / 'xlsx'


class StyleExcel(BaseModel):
    plage: str
    fond: str = ''
    couleur: str = '202020'
    gras: bool = False
    taille: int = Field(default=11, ge=8, le=40)
    format_nombre: str = Field(default='', max_length=100)
    alignement: Literal['left', 'center', 'right'] = 'left'
    retour_ligne: bool = True


class FormuleExcel(BaseModel):
    cellule: str
    formule: str = Field(max_length=2000)


class GraphiqueExcel(BaseModel):
    type: Literal['barres', 'courbe'] = 'barres'
    titre: str = Field(default='', max_length=200)
    plage: str
    ancrage: str = 'H2'


class ImageExcel(BaseModel):
    image_base64: str = Field(default='', max_length=5400000)
    svg: str = Field(default='', max_length=300000)
    ancrage: str = 'H2'
    largeur_px: int = Field(default=600, ge=20, le=2000)


class FeuilleExcel(BaseModel):
    images: list[ImageExcel] = Field(default_factory=list, max_length=20)
    nom: str = Field(min_length=1, max_length=31)
    lignes: list[list[str | int | float | bool | None]] = Field(max_length=5000)
    entete: bool = True
    couleur_entete: str = '176B87'
    filtre: bool = True
    figer: str = 'A2'
    largeurs: dict[str, float] = Field(default_factory=dict)
    styles: list[StyleExcel] = Field(default_factory=list, max_length=100)
    formules: list[FormuleExcel] = Field(default_factory=list, max_length=5000)
    graphiques: list[GraphiqueExcel] = Field(default_factory=list, max_length=10)


def _couleur(valeur):
    valeur = valeur.lstrip('#')
    if not re.fullmatch(r'[0-9A-Fa-f]{6}', valeur):
        raise ValueError('Couleur attendue : six caractères hexadécimaux.')
    return valeur


def _cellule(valeur):
    if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]{0,3}', valeur):
        raise ValueError(f'Adresse de cellule non autorisée : {valeur}')
    _plage(valeur)
    return valeur


def _plage(valeur):
    if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]{0,3}(?::[A-Z]{1,3}[1-9][0-9]{0,3})?', valeur):
        raise ValueError(f'Plage non autorisée : {valeur}')
    a, b, c, d = range_boundaries(valeur)
    if not (1 <= a <= c <= 100 and 1 <= b <= d <= 5000):
        raise ValueError('Plage limitée à 100 colonnes et 5000 lignes.')
    if (c-a+1)*(d-b+1) > 100000:
        raise ValueError('Plage trop volumineuse.')
    return a, b, c, d


@mcp.tool()
def creer_excel(
    feuilles: list[FeuilleExcel],
    nom: str = 'classeur',
    profil: str = 'libre',
    modele: str = '',
) -> dict:
    """Crée un véritable XLSX avec plusieurs feuilles et un lien temporaire.

    lignes : valeurs littérales. Les chaînes commençant par = restent du texte.
    formules : liste séparée {cellule, formule}. Employer les noms de fonctions
    anglais et les virgules. Formules simples autorisées : calculs, références,
    fonctions usuelles ; pas de liens externes ni de fonctions réseau.
    Les formules sont enregistrées mais PAS calculées sur le serveur.
    styles : plages A1:D10, couleurs hexadécimales, format_nombre, alignement.
    graphiques : plage rectangulaire comprenant les en-têtes en première ligne,
    catégories dans la première colonne et séries numériques dans les suivantes.
    Limites : 20 feuilles, 100 colonnes, 5000 lignes par feuille,
    100000 cellules de données au total. Sources dans documents_generes/source.
    profil et modele sont seulement des métadonnées déclaratives.
    images : liste {image_base64 OU svg, ancrage, largeur_px} par feuille.
    Base64 PNG/JPEG brut ; SVG autonome converti en PNG.
    Aucun aperçu visuel n'est effectué.
    """
    debut = perf_counter()
    if not re.fullmatch(r'[A-Za-z0-9_-]{1,60}', nom):
        raise ValueError('Nom : 1 à 60 lettres sans accents, chiffres, tirets ou underscores.')
    if not 1 <= len(feuilles) <= 20:
        raise ValueError('Prévoir entre 1 et 20 feuilles.')
    if len(profil) > 80 or len(modele) > 200:
        raise ValueError('Métadonnées trop longues.')
    noms = [f.nom.casefold() for f in feuilles]
    if len(noms) != len(set(noms)):
        raise ValueError('Noms de feuilles dupliqués.')
    if sum(len(l) for f in feuilles for l in f.lignes) > 100000:
        raise ValueError('Plus de 100000 cellules de données.')
    DOSSIER.mkdir(exist_ok=True)
    SOURCES.mkdir(parents=True, exist_ok=True)
    if len(list(DOSSIER.glob('*.xlsx'))) >= 200:
        raise ValueError('Limite de 200 classeurs atteinte.')
    source = {'profil': profil, 'modele': modele.strip() or 'non renseigne',
              'feuilles': [f.model_dump() for f in feuilles]}
    # Éviter notamment les chaînes géantes et les nombres non JSON.
    brut = json.dumps(source, ensure_ascii=False, allow_nan=False)
    if len(brut.encode('utf-8')) > 3000000:
        raise ValueError('Source limitée à 3 Mo.')
    wb = Workbook()
    wb.remove(wb.active)
    for f in feuilles:
        if re.search(r'[\\/*?:\[\]]', f.nom) or f.nom.startswith("'") or f.nom.endswith("'"):
            raise ValueError('Nom de feuille invalide.')
        if not f.lignes or max((len(l) for l in f.lignes), default=0) > 100:
            raise ValueError('Feuille vide ou plus de 100 colonnes.')
        ws = wb.create_sheet(f.nom)
        for i, ligne in enumerate(f.lignes, 1):
            for j, valeur in enumerate(ligne, 1):
                if isinstance(valeur, str) and len(valeur) > 32767:
                    raise ValueError('Texte de cellule trop long.')
                cellule = ws.cell(i, j, valeur)
                if isinstance(valeur, str):
                    cellule.data_type = 's'
                cellule.alignment = Alignment(vertical='top', wrap_text=True)
        if f.entete:
            couleur = _couleur(f.couleur_entete)
            for cellule in ws[1]:
                cellule.fill = PatternFill('solid', fgColor=couleur)
                cellule.font = Font(bold=True, color='FFFFFF')
        if f.filtre:
            ws.auto_filter.ref = ws.dimensions
        if f.figer:
            ws.freeze_panes = _cellule(f.figer)
        for j in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(j)].width = 22
        for colonne, largeur in f.largeurs.items():
            _cellule(colonne + '1')
            if not math.isfinite(largeur) or not 3 <= largeur <= 100:
                raise ValueError('Largeur attendue : 3 à 100.')
            ws.column_dimensions[colonne].width = largeur
        for s in f.styles:
            a, b, c, d = _plage(s.plage)
            couleur = _couleur(s.couleur)
            fond = _couleur(s.fond) if s.fond else None
            for ligne in ws.iter_rows(min_row=b, max_row=d, min_col=a, max_col=c):
                for cellule in ligne:
                    cellule.font = Font(size=s.taille, bold=s.gras, color=couleur)
                    cellule.alignment = Alignment(horizontal=s.alignement, vertical='top', wrap_text=s.retour_ligne)
                    if fond:
                        cellule.fill = PatternFill('solid', fgColor=fond)
                    if s.format_nombre:
                        cellule.number_format = s.format_nombre
        autorisees = {'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'IF',
                      'AND', 'OR', 'NOT', 'ROUND', 'ABS', 'SUMIF', 'COUNTIF', 'IFERROR'}
        for fml in f.formules:
            _cellule(fml.cellule)
            expression = fml.formule
            if not expression.startswith('=') or re.search(r"[\[\]{}|!\\]|://", expression):
                raise ValueError('Formule non autorisée ; pas de références externes ou interfeuilles.')
            fonctions = re.findall(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(', expression)
            if any(x.upper() not in autorisees for x in fonctions):
                raise ValueError('Fonction non autorisée dans une formule.')
            ws[fml.cellule] = expression
        for illustration in f.images:
            from openpyxl.drawing.image import Image as XLImage
            _cellule(illustration.ancrage)
            png = image_png(illustration.image_base64, illustration.svg)
            image = XLImage(BytesIO(png))
            image.height = round(image.height * illustration.largeur_px / image.width)
            image.width = illustration.largeur_px
            ws.add_image(image, illustration.ancrage)
        for g in f.graphiques:
            a, b, c, d = _plage(g.plage)
            if c <= a or d <= b:
                raise ValueError('Graphique : au moins deux colonnes et deux lignes.')
            _cellule(g.ancrage)
            chart = BarChart() if g.type == 'barres' else LineChart()
            chart.title = g.titre
            chart.add_data(Reference(ws, min_col=a+1, max_col=c, min_row=b, max_row=d), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=a, min_row=b+1, max_row=d))
            ws.add_chart(chart, g.ancrage)
    tampon = BytesIO()
    wb.save(tampon)
    contenu = tampon.getvalue()
    if len(contenu) > 12000000:
        raise ValueError('Classeur trop volumineux.')
    creation = datetime.now(ZoneInfo('Europe/Paris'))
    source['cree_le'] = creation.isoformat(timespec='seconds')
    source['fuseau_horaire'] = 'Europe/Paris'
    brut = json.dumps(source, ensure_ascii=False, indent=2, allow_nan=False)
    base = f'{nom}_{creation:%Y-%m-%d_%Hh%Mm%Ss}'
    version = 1
    while True:
        suffixe = '' if version == 1 else f'-v{version:02d}'
        fichier = DOSSIER / f'{base}{suffixe}.xlsx'
        chemin_source = SOURCES / f'{fichier.stem}.xlsx.source.json'
        crees = []
        try:
            with fichier.open('xb') as flux:
                crees.append(fichier)
                fichier.chmod(0o600)
                flux.write(contenu)
            with chemin_source.open('x', encoding='utf-8') as flux:
                crees.append(chemin_source)
                chemin_source.chmod(0o600)
                flux.write(brut)
        except FileExistsError:
            for p in crees:
                p.unlink(missing_ok=True)
            version += 1
            continue
        except Exception:
            for p in crees:
                p.unlink(missing_ok=True)
            raise
        break
    from telechargements import creer_lien
    try:
        lien = creer_lien(f'documents_generes/xlsx/{fichier.name}')
    except Exception:
        lien = {'statut': 'indisponible', 'message': 'Classeur créé ; réessayer avec lien_telechargement.'}
    return {'statut': 'cree', 'fichier': f'documents_generes/xlsx/{fichier.name}',
            'source': f'documents_generes/source/xlsx/{chemin_source.name}',
            'feuilles': wb.sheetnames, 'taille_octets': len(contenu),
            'telechargement': lien, 'rendu_visuel_verifie': False,
            'formules_calculees_sur_serveur': False,
            'mesures': {'duree_creation_secondes': round(perf_counter()-debut, 3)}}
