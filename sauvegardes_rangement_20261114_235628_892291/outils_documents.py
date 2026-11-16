from images_documents import image_png
import re
from pathlib import Path
from typing import Literal
from datetime import datetime
from zoneinfo import ZoneInfo

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pydantic import BaseModel, Field

from serveur_mcp import mcp


DOSSIER = Path(__file__).resolve().parent / "documents_generes"
DOSSIER.mkdir(exist_ok=True)


class BlocDocument(BaseModel):
    type: Literal[
        "titre", "paragraphe", "puces",
        "tableau", "encadre", "saut_page", "image", "svg"
    ]
    image_base64: str = Field(default='', max_length=5400000)
    svg: str = Field(default='', max_length=300000)
    largeur_cm: float = Field(default=12, ge=0.5, le=19)
    texte: str = Field(default="", max_length=20000)
    niveau: int = Field(default=1, ge=1, le=3)
    alignement: Literal[
        "gauche", "centre", "droite", "justifie"
    ] = "gauche"
    elements: list[str] = Field(default_factory=list, max_length=80)
    lignes: list[list[str]] = Field(default_factory=list, max_length=80)


def ajouter_texte(paragraphe, texte):
    """Interprète **texte** comme du gras."""
    for morceau in re.split(r"(\*\*.*?\*\*)", texte, flags=re.S):
        if morceau.startswith("**") and morceau.endswith("**"):
            paragraphe.add_run(morceau[2:-2]).bold = True
        else:
            paragraphe.add_run(morceau)


def appliquer_fond(cellule, couleur):
    fond = OxmlElement("w:shd")
    fond.set(qn("w:fill"), couleur)
    cellule._tc.get_or_add_tcPr().append(fond)


@mcp.tool()
def creer_document(
    titre: str,
    blocs: list[BlocDocument],
    nom: str = "document",
    sous_titre: str = "",
    couleur: str = "176B87",
    police: str = "Arial",
    taille_texte: int = 11,
    marge_cm: float = 2.0,
) -> dict:
    """Crée un fichier DOCX mis en page, destiné notamment à Pages.

    Types de blocs :
    - titre : texte et niveau de 1 à 3.
    - paragraphe : texte et alignement.
    - puces : elements, liste de textes.
    - tableau : lignes, liste de lignes de cellules.
      La première ligne sert d'en-tête.
    - encadre : texte dans un encadré coloré.
    - saut_page : commence une nouvelle page.
    - image : image_base64 PNG/JPEG brut, largeur_cm, alignement, texte légende.
    - svg : svg XML autonome, largeur_cm ; converti en PNG.

    Dans les paragraphes, puces et encadrés, **texte** donne du gras.
    couleur : code hexadécimal de six caractères, sans #.
    Le résultat donne un chemin dans Codespaces, PAS un lien web.
    Aucun accès aux fichiers du Mac ni aux commandes système.
    """
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,60}", nom):
        raise ValueError(
            "nom : 1 à 60 lettres sans accents, chiffres, tirets ou underscores."
        )
    couleur = couleur.lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", couleur):
        raise ValueError("La couleur doit être un code hexadécimal à 6 caractères.")
    if not (8 <= taille_texte <= 20 and 1 <= marge_cm <= 4):
        raise ValueError("Taille : 8 à 20 points. Marges : 1 à 4 cm.")
    if len(blocs) > 100 or len(titre) > 300 or len(sous_titre) > 1000:
        raise ValueError("Document trop volumineux pour cet outil.")
    if len(police) > 80:
        raise ValueError("Nom de police trop long.")
    if len(list(DOSSIER.glob("*.docx"))) >= 200:
        raise ValueError("Limite de 200 documents atteinte. Fais du ménage.")

    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(marge_cm)
    section.left_margin = section.right_margin = Cm(marge_cm)

    normal = doc.styles["Normal"]
    normal.font.name = police
    normal.font.size = Pt(taille_texte)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.15

    for style in ("Title", "Heading 1", "Heading 2", "Heading 3"):
        doc.styles[style].font.name = police
        doc.styles[style].font.color.rgb = RGBColor.from_string(couleur)

    doc.add_heading(titre, 0)
    if sous_titre:
        doc.add_paragraph(sous_titre, style="Subtitle")

    alignements = {
        "gauche": WD_ALIGN_PARAGRAPH.LEFT,
        "centre": WD_ALIGN_PARAGRAPH.CENTER,
        "droite": WD_ALIGN_PARAGRAPH.RIGHT,
        "justifie": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }

    for bloc in blocs:
        if bloc.type == "titre":
            doc.add_heading(bloc.texte, bloc.niveau)

        elif bloc.type == "paragraphe":
            p = doc.add_paragraph()
            p.alignment = alignements[bloc.alignement]
            ajouter_texte(p, bloc.texte)

        elif bloc.type == "puces":
            for element in bloc.elements:
                p = doc.add_paragraph(style="List Bullet")
                ajouter_texte(p, element)

        elif bloc.type == "tableau":
            lignes = bloc.lignes
            if not lignes or not 1 <= len(lignes[0]) <= 8:
                raise ValueError("Un tableau doit avoir de 1 à 8 colonnes.")
            if any(len(ligne) != len(lignes[0]) for ligne in lignes):
                raise ValueError("Toutes les lignes doivent avoir la même longueur.")

            tableau = doc.add_table(rows=0, cols=len(lignes[0]))
            tableau.style = "Table Grid"
            for index, ligne in enumerate(lignes):
                cellules = tableau.add_row().cells
                for cellule, texte in zip(cellules, ligne):
                    cellule.text = texte
                    if index == 0:
                        appliquer_fond(cellule, couleur)
                        for p in cellule.paragraphs:
                            for run in p.runs:
                                run.bold = True
                                run.font.color.rgb = RGBColor(255, 255, 255)
            doc.add_paragraph()

        elif bloc.type == "encadre":
            tableau = doc.add_table(rows=1, cols=1)
            tableau.style = "Table Grid"
            cellule = tableau.cell(0, 0)
            appliquer_fond(cellule, "EDF4F8")
            ajouter_texte(cellule.paragraphs[0], bloc.texte)
            doc.add_paragraph()

        elif bloc.type in ("image", "svg"):
            from io import BytesIO
            png = image_png(bloc.image_base64, bloc.svg)
            p = doc.add_paragraph()
            p.alignment = alignements[bloc.alignement]
            p.add_run().add_picture(
                BytesIO(png), width=Cm(min(bloc.largeur_cm, 21 - 2*marge_cm))
            )
            if bloc.texte:
                ajouter_texte(doc.add_paragraph(), bloc.texte)

        elif bloc.type == "saut_page":
            doc.add_page_break()

    creation = datetime.now(ZoneInfo("Europe/Paris"))
    base = f"{nom}_{creation:%Y-%m-%d_%Hh%Mm%Ss}"
    version = 1

    while True:
        suffixe = "" if version == 1 else f"-v{version:02d}"
        fichier = DOSSIER / f"{base}{suffixe}.docx"

        try:
            sortie = fichier.open("xb")
        except FileExistsError:
            version += 1
            continue

        try:
            with sortie:
                doc.save(sortie)
        except Exception:
            fichier.unlink(missing_ok=True)
            raise

        break


    return {
        "statut": "cree",
        "fichier": f"documents_generes/{fichier.name}",
        "taille_octets": fichier.stat().st_size,
        "recuperation": "Télécharger depuis l'explorateur de Codespaces.",
        "rendu_visuel_verifie": False,
    }


@mcp.tool()
def lister_documents() -> list[str]:
    """Liste les DOCX du dossier documents_generes, sans lire leur contenu."""
    return sorted(
        f"documents_generes/{f.name}"
        for f in DOSSIER.glob("*.docx")
        if f.is_file()
    )
