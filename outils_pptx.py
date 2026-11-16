from images_documents import image_png
"""Création de présentations PPTX éditables, à composition libre."""
import json
import math
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from time import perf_counter
from typing import Literal
from zoneinfo import ZoneInfo

from pydantic import BaseModel, Field
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm, Pt

from serveur_mcp import mcp

from rangement_documents import dossier_format, trouver_document
DOSSIER = dossier_format("pptx")
SOURCES = DOSSIER.parent / 'source' / 'pptx'
DOSSIER.mkdir(exist_ok=True)
SOURCES.mkdir(parents=True, exist_ok=True)


def couleur(valeur: str) -> RGBColor:
    if not re.fullmatch(r'[0-9a-fA-F]{6}', valeur):
        raise ValueError('Couleur : six caractères hexadécimaux, sans #.')
    return RGBColor.from_string(valeur)


class SeriePPTX(BaseModel):
    nom: str = Field(max_length=150)
    valeurs: list[float] = Field(min_length=1, max_length=200)


class ElementPPTX(BaseModel):
    type: Literal['texte', 'rectangle', 'arrondi', 'ellipse', 'fleche', 'ligne', 'tableau', 'graphique', 'image', 'svg']
    image_base64: str = Field(default='', max_length=5400000)
    svg: str = Field(default='', max_length=300000)
    x: float = Field(default=1, ge=0, le=60)
    y: float = Field(default=1, ge=0, le=60)
    largeur: float = Field(default=10, gt=0, le=60)
    hauteur: float = Field(default=2, gt=0, le=60)
    texte: str = Field(default='', max_length=12000)
    police: str = Field(default='Arial', min_length=1, max_length=80)
    taille: float = Field(default=20, ge=8, le=96)
    gras: bool = False
    italique: bool = False
    couleur_texte: str = '202830'
    fond: str = ''
    contour: str = ''
    epaisseur: float = Field(default=1, ge=0, le=10)
    alignement: Literal['gauche', 'centre', 'droite'] = 'gauche'
    lignes: list[list[str]] = Field(default_factory=list, max_length=40)
    entete: bool = True
    couleur_entete: str = '245675'
    categories: list[str] = Field(default_factory=list, max_length=200)
    series: list[SeriePPTX] = Field(default_factory=list, max_length=12)
    graphique: Literal['colonnes', 'barres', 'courbes', 'secteurs'] = 'colonnes'


class DiapositivePPTX(BaseModel):
    fond: str = 'FFFFFF'
    notes: str = Field(default='', max_length=15000)
    elements: list[ElementPPTX] = Field(default_factory=list, max_length=100)


def appliquer_texte(cadre, texte: str, e: ElementPPTX, teinte=None, gras=None):
    cadre.clear()
    cadre.word_wrap = True
    cadre.margin_left = Cm(0.12)
    cadre.margin_right = Cm(0.12)
    cadre.margin_top = Cm(0.08)
    cadre.margin_bottom = Cm(0.08)
    alignements = {'gauche': PP_ALIGN.LEFT, 'centre': PP_ALIGN.CENTER, 'droite': PP_ALIGN.RIGHT}
    for i, ligne in enumerate(texte.split('\n')):
        p = cadre.paragraphs[0] if i == 0 else cadre.add_paragraph()
        p.alignment = alignements[e.alignement]
        p.space_after = Pt(4)
        r = p.add_run()
        r.text = ligne
        r.font.name = e.police
        r.font.size = Pt(e.taille)
        r.font.bold = e.gras if gras is None else gras
        r.font.italic = e.italique
        r.font.color.rgb = couleur(teinte or e.couleur_texte)


def verifier_element(e: ElementPPTX):
    for v in (e.x, e.y, e.largeur, e.hauteur, e.taille, e.epaisseur):
        if not math.isfinite(v):
            raise ValueError('Les dimensions doivent être finies.')
    for v in (e.couleur_texte, e.fond, e.contour, e.couleur_entete):
        if v:
            couleur(v)
    if e.type == 'tableau':
        if not e.lignes or not 1 <= len(e.lignes[0]) <= 12:
            raise ValueError('Tableau : 1 à 12 colonnes et au moins une ligne.')
        if any(len(ligne) != len(e.lignes[0]) for ligne in e.lignes):
            raise ValueError('Toutes les lignes du tableau doivent avoir la même longueur.')
        if sum(len(cellule) for ligne in e.lignes for cellule in ligne) > 30000:
            raise ValueError('Contenu du tableau trop volumineux.')
    if e.type == 'graphique':
        if not e.categories or not e.series:
            raise ValueError('Graphique : categories et series obligatoires.')
        if e.graphique == 'secteurs' and len(e.series) != 1:
            raise ValueError('Un graphique en secteurs accepte une seule série.')
        for serie in e.series:
            if len(serie.valeurs) != len(e.categories):
                raise ValueError('Chaque série doit avoir une valeur par catégorie.')
            if not all(math.isfinite(v) for v in serie.valeurs):
                raise ValueError('Les valeurs des séries doivent être finies.')


@mcp.tool()
def creer_presentation(
    titre: str,
    diapositives: list[DiapositivePPTX],
    nom: str = 'presentation',
    profil: Literal['chatgpt', 'claude', 'gemini', 'libre'] = 'libre',
    modele: str = '',
    largeur_cm: float = 33.867,
    hauteur_cm: float = 19.05,
) -> dict:
    """Crée un vrai PPTX éditable à composition libre, sans thème imposé.

    Coordonnées x, y, largeur et hauteur en CENTIMÈTRES ; taille en POINTS.
    Format par défaut 16:9. Chaque diapositive reçoit son fond et ses éléments
    dans l'ordre d'empilement (premier derrière, dernier devant).
    Aucun titre n'est ajouté automatiquement : composer tous les éléments.
    Types : texte, rectangle, arrondi, ellipse, fleche (vers la droite), ligne
    (du coin haut-gauche au coin bas-droit), tableau, graphique.
    Couleurs : six chiffres hexadécimaux sans # ; fond/contour vides = absents.
    Texte brut avec retours à la ligne ; Markdown et HTML non interprétés.
    Tableaux : lignes rectangulaires ; entete colore la première ligne.
    Graphiques : categories, series [{nom, valeurs}], graphique parmi colonnes,
    barres, courbes, secteurs. Images : type image + image_base64 PNG/JPEG brut ; type svg + svg XML.
    Le SVG est converti en PNG ; proportions conservées dans le cadre.
    Pas d'animations ou aperçu inclus.
    Choisir librement le nombre de diapositives adapté à la demande ; respecter
    le nombre demandé s'il est précisé. profil est descriptif, pas un gabarit.
    modele : identité fiable du modèle actif, sinon vide ; source seulement.
    Le résultat inclut le chemin et un lien de téléchargement si disponible.
    La création ne constitue pas un contrôle visuel ou une preuve de lisibilité.
    """
    debut = perf_counter()
    if not re.fullmatch(r'[A-Za-z0-9_-]{1,60}', nom):
        raise ValueError('Nom : 1 à 60 lettres sans accents, chiffres, tirets ou underscores.')
    if not 1 <= len(diapositives) <= 60:
        raise ValueError('La présentation doit contenir de 1 à 60 diapositives.')
    if len(titre) > 300 or len(modele) > 150:
        raise ValueError('Titre ou identifiant du modèle trop long.')
    if not (math.isfinite(largeur_cm) and math.isfinite(hauteur_cm)
            and 10 <= largeur_cm <= 60 and 10 <= hauteur_cm <= 60):
        raise ValueError('Dimensions : entre 10 et 60 cm.')
    if sum(len(d.elements) for d in diapositives) > 1000:
        raise ValueError('Limite de 1000 éléments par présentation.')
    if sum(1 for _ in DOSSIER.glob('*.pptx')) >= 200:
        raise ValueError('Limite de 200 PPTX atteinte.')

    donnees = {
        'profil': profil, 'modele': modele.strip() or 'non renseigne',
        'titre': titre, 'largeur_cm': largeur_cm, 'hauteur_cm': hauteur_cm,
        'diapositives': [d.model_dump() for d in diapositives],
    }
    if len(json.dumps(donnees, ensure_ascii=False).encode('utf-8')) > 2_000_000:
        raise ValueError('Source limitée à 2 Mo.')
    for d in diapositives:
        couleur(d.fond)
        for e in d.elements:
            verifier_element(e)

    prs = Presentation()
    prs.slide_width = Cm(largeur_cm)
    prs.slide_height = Cm(hauteur_cm)
    prs.core_properties.title = titre
    alertes = []
    formes = {
        'rectangle': MSO_SHAPE.RECTANGLE,
        'arrondi': MSO_SHAPE.ROUNDED_RECTANGLE,
        'ellipse': MSO_SHAPE.OVAL,
        'fleche': MSO_SHAPE.RIGHT_ARROW,
    }
    types_graphiques = {
        'colonnes': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'barres': XL_CHART_TYPE.BAR_CLUSTERED,
        'courbes': XL_CHART_TYPE.LINE,
        'secteurs': XL_CHART_TYPE.PIE,
    }
    for numero, d in enumerate(diapositives, 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = couleur(d.fond)
        if d.notes:
            slide.notes_slide.notes_text_frame.text = d.notes
        for index, e in enumerate(d.elements, 1):
            if e.x + e.largeur > largeur_cm + 0.01 or e.y + e.hauteur > hauteur_cm + 0.01:
                alertes.append(f'Diapositive {numero}, élément {index} : dépassement du cadre.')
            x, y, w, h = map(Cm, (e.x, e.y, e.largeur, e.hauteur))
            if e.type in ('image', 'svg'):
                png = image_png(e.image_base64, e.svg)
                from PIL import Image as PILImage
                with PILImage.open(BytesIO(png)) as im:
                    facteur = min(w / im.width, h / im.height)
                    iw, ih = int(im.width * facteur), int(im.height * facteur)
                slide.shapes.add_picture(BytesIO(png), x + (w-iw)//2, y + (h-ih)//2, width=iw, height=ih)
                continue
            if e.type == 'tableau':
                table = slide.shapes.add_table(len(e.lignes), len(e.lignes[0]), x, y, w, h).table
                for r, ligne in enumerate(e.lignes):
                    for c, valeur in enumerate(ligne):
                        cellule = table.cell(r, c)
                        if r == 0 and e.entete:
                            cellule.fill.solid()
                            cellule.fill.fore_color.rgb = couleur(e.couleur_entete)
                            appliquer_texte(cellule.text_frame, valeur, e, 'FFFFFF', True)
                        else:
                            if e.fond:
                                cellule.fill.solid()
                                cellule.fill.fore_color.rgb = couleur(e.fond)
                            appliquer_texte(cellule.text_frame, valeur, e)
                continue
            if e.type == 'graphique':
                data = CategoryChartData()
                data.categories = e.categories
                for serie in e.series:
                    data.add_series(serie.nom, serie.valeurs)
                chart = slide.shapes.add_chart(types_graphiques[e.graphique], x, y, w, h, data).chart
                chart.has_legend = len(e.series) > 1 or e.graphique == 'secteurs'
                continue
            if e.type == 'ligne':
                shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x, y, x + w, y + h)
                shape.line.color.rgb = couleur(e.contour or e.couleur_texte)
                shape.line.width = Pt(e.epaisseur)
                continue
            if e.type == 'texte':
                shape = slide.shapes.add_textbox(x, y, w, h)
            else:
                shape = slide.shapes.add_shape(formes[e.type], x, y, w, h)
            if e.fond:
                shape.fill.solid()
                shape.fill.fore_color.rgb = couleur(e.fond)
            else:
                shape.fill.background()
            if e.contour:
                shape.line.color.rgb = couleur(e.contour)
                shape.line.width = Pt(e.epaisseur)
            else:
                shape.line.fill.background()
            appliquer_texte(shape.text_frame, e.texte, e)

    buffer = BytesIO()
    prs.save(buffer)
    contenu = buffer.getvalue()
    if len(contenu) > 20_000_000:
        raise ValueError('Fichier limité à 20 Mo.')
    # Réouverture technique : ne vérifie pas le rendu visuel.
    nombre = len(Presentation(BytesIO(contenu)).slides)
    creation = datetime.now(ZoneInfo('Europe/Paris'))
    donnees.update({
        'cree_le': creation.isoformat(timespec='seconds'),
        'fuseau_horaire': 'Europe/Paris',
        'nombre_diapositives': nombre,
    })
    source_json = json.dumps(donnees, ensure_ascii=False, indent=2)
    base = f'{nom}_{creation:%Y-%m-%d_%Hh%Mm%Ss}'
    version = 1
    while True:
        suffixe = '' if version == 1 else f'-v{version:02d}'
        fichier = DOSSIER / f'{base}{suffixe}.pptx'
        source = SOURCES / f'{base}{suffixe}.pptx.source.json'
        crees = []
        try:
            with source.open('x', encoding='utf-8') as f:
                crees.append(source)
                source.chmod(0o600)
                f.write(source_json)
            with fichier.open('xb') as f:
                crees.append(fichier)
                fichier.chmod(0o600)
                f.write(contenu)
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
    try:
        from telechargements import creer_lien
        lien = creer_lien(f'documents_generes/pptx/{fichier.name}')
    except Exception:
        lien = {'statut': 'indisponible', 'message': 'Utiliser lien_telechargement pour réessayer.'}
    return {
        'statut': 'cree',
        'fichier': f'documents_generes/pptx/{fichier.name}',
        'source': f'documents_generes/source/pptx/{source.name}',
        'format': 'pptx', 'profil': profil,
        'nombre_diapositives': nombre,
        'taille_octets': len(contenu),
        'mesures': {'duree_creation_secondes': round(perf_counter() - debut, 3)},
        'telechargement': lien,
        'alertes': alertes,
        'controle_technique': 'PPTX réouvert et nombre de diapositives vérifié. Dépassements géométriques recherchés.',
        'rendu_visuel_verifie': False,
        'limites_controle': 'Ni débordements du texte, ni chevauchements, ni esthétique vérifiés.',
    }
