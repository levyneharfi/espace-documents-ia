"""Contrôle automatique de la mise en page d'un PDF, sans vision (texte uniquement)."""
import re
import subprocess
import tempfile
from html import unescape
from pathlib import Path

from PIL import Image

from serveur_mcp import mcp

DOSSIER = (Path(__file__).resolve().parent / "documents_generes").resolve()
MM = 72 / 25.4            # points PDF par millimètre
DPI = 50                  # résolution du rendu d'analyse
PX_MM = DPI / 25.4        # pixels par millimètre à cette résolution
NIVEAUX = " .:-=+*#%@"
MAX_CARTES = 30


def _trouver(fichier: str) -> Path:
    from rangement_documents import trouver_document
    return trouver_document(fichier, {'.pdf'})


def _texte_par_page(chemin: Path):
    r = subprocess.run(["pdftotext", "-bbox", str(chemin), "-"],
                       capture_output=True, text=True, timeout=60, check=True)
    pages = []
    for p in re.finditer(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', r.stdout, re.S):
        mots = []
        for w in re.finditer(r'<word xMin="([-\d.]+)" yMin="([-\d.]+)" xMax="([-\d.]+)" yMax="([-\d.]+)">(.*?)</word>',
                             p.group(3), re.S):
            x0, y0, x1, y1 = map(float, w.groups()[:4])
            mots.append((x0, y0, x1, y1, unescape(w.group(5))))
        pages.append({"largeur": float(p.group(1)), "hauteur": float(p.group(2)), "mots": mots})
    return pages


def _lignes(mots):
    lignes = []
    for x0, y0, x1, y1, t in sorted(mots, key=lambda m: (round(m[1]), m[0])):
        if lignes and abs(lignes[-1]["y"] - y0) < 3:
            lignes[-1]["mots"].append((x0, t))
        else:
            lignes.append({"y": y0, "mots": [(x0, t)]})
    return [" ".join(t for _, t in sorted(l["mots"])) for l in lignes]


def _chevauchements(mots, limite=5):
    trouves = []
    if len(mots) > 1500:
        return trouves
    for i, a in enumerate(mots):
        for b in mots[i + 1:]:
            ix = min(a[2], b[2]) - max(a[0], b[0])
            iy = min(a[3], b[3]) - max(a[1], b[1])
            if ix <= 0 or iy <= 0:
                continue
            petite = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
            if petite > 0 and ix * iy > 0.3 * petite:
                trouves.append(f"« {a[4]} » / « {b[4]} »")
                if len(trouves) >= limite:
                    return trouves
    return trouves


def _analyse_image(png: Path):
    with Image.open(png) as img:
        gris = img.convert("L")
    w, h = gris.size
    hist = gris.histogram()
    fond = hist.index(max(hist))  # couleur la plus fréquente = fond de page
    masque = gris.point(lambda p: 255 if abs(p - fond) > 20 else 0)
    cols = 40
    rows = max(1, round(cols * h / w / 2))  # un caractère est environ deux fois plus haut que large
    petit = masque.resize((cols, rows), Image.Resampling.BOX)
    carte = []
    for y in range(rows):
        ligne = "".join(NIVEAUX[min(len(NIVEAUX) - 1, petit.getpixel((x, y)) * len(NIVEAUX) // 256)]
                        for x in range(cols))
        carte.append("|" + ligne + "|")
    return {"w": w, "h": h, "bbox": masque.getbbox(),
            "remplissage": masque.histogram()[255] / (w * h), "carte": carte}


def analyser_pdf(fichier: str) -> dict:
    chemin = _trouver(fichier)
    textes = _texte_par_page(chemin)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", str(DPI), "-png", str(chemin), f"{tmp}/p"],
                       check=True, timeout=120, capture_output=True)
        images = sorted(Path(tmp).glob("p-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
        rendus = [_analyse_image(p) for p in images]

    n = len(rendus)
    alertes, details = [], []
    for i, r in enumerate(rendus, start=1):
        pt = textes[i - 1] if i - 1 < len(textes) else {"largeur": 0, "hauteur": 0, "mots": []}
        W, H, mots = pt["largeur"], pt["hauteur"], pt["mots"]
        d = {"page": i, "nombre_de_mots": len(mots), "encre_pourcentage": round(r["remplissage"] * 100, 1)}

        if r["bbox"] is None:
            alertes.append(f"Page {i} : page blanche.")
        else:
            g, t, dr, b = r["bbox"]
            marges = {"haut": round(t / PX_MM), "bas": round((r["h"] - b) / PX_MM),
                      "gauche": round(g / PX_MM), "droite": round((r["w"] - dr) / PX_MM)}
            d["marges_reelles_mm"] = marges
            d["contenu_jusqu_a_pourcentage_hauteur"] = round(b / r["h"] * 100)
            if min(marges.values()) < 3:
                alertes.append(f"Page {i} : des éléments touchent le bord (moins de 3 mm) — "
                               "normal pour un fond ou un bandeau pleine page, sinon débordement probable.")
            if i == n and n > 1 and b / r["h"] < 0.25:
                alertes.append(f"Page {i} : dernière page presque vide (contenu sur {round(b / r['h'] * 100)} % "
                               "de la hauteur) — probable débordement de quelques lignes.")

        if W and H:
            hors = [m[4] for m in mots if m[0] < 0 or m[1] < 0 or m[2] > W or m[3] > H]
            bord = [m[4] for m in mots if m[4] not in hors and
                    (m[0] < 8 * MM or m[1] < 8 * MM or m[2] > W - 8 * MM or m[3] > H - 8 * MM)]
            if hors:
                alertes.append(f"Page {i} : texte hors de la page (coupé) : {', '.join(hors[:5])}")
            if bord:
                alertes.append(f"Page {i} : texte à moins de 8 mm du bord (pied de page ?) : {', '.join(bord[:5])}")
        chev = _chevauchements(mots)
        if chev:
            alertes.append(f"Page {i} : textes qui se chevauchent : {'; '.join(chev)}")

        lignes = _lignes(mots)
        d["premieres_lignes"] = lignes[:2]
        d["dernieres_lignes"] = lignes[-2:]
        if i <= MAX_CARTES:
            d["carte_densite"] = r["carte"]
        details.append(d)

    fmt = f"{round(textes[0]['largeur'] / MM)} × {round(textes[0]['hauteur'] / MM)} mm" if textes else "inconnu"
    return {
        "fichier": chemin.name,
        "nombre_de_pages": n,
        "format": fmt,
        "alertes": alertes or ["Aucune alerte automatique."],
        "pages": details,
        "nature_du_controle": ("Mesures automatiques sur le PDF rendu (texte extrait et pixels). "
                               "Ce n'est pas un contrôle visuel : esthétique, couleurs et lisibilité "
                               "réelle ne sont pas évaluées."),
    }


@mcp.tool()
def controle_pdf(fichier: str) -> dict:
    """Contrôle automatique de la mise en page d'un PDF de documents_generes, en texte (sans image) :
    pages, format, marges réelles, texte coupé ou trop près du bord, chevauchements, pages vides,
    premières et dernières lignes de chaque page, et carte de densité (une ligne de caractères
    par bande horizontale de la page ; ' ' = vide, '@' = très rempli).
    Donner le nom ou le chemin exact renvoyé par creer_pdf. À utiliser quand apercu_pdf
    ne transmet pas d'image au modèle."""
    return analyser_pdf(fichier)
