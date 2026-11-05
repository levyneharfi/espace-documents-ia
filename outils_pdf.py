import json
import re
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Lock
from typing import Literal
from datetime import datetime
from zoneinfo import ZoneInfo

from mcp.server.fastmcp import Image

from serveur_mcp import mcp


DOSSIER = Path(__file__).resolve().parent / "documents_generes"
DOSSIER.mkdir(exist_ok=True)

VERROU_RENDU = Lock()


# Programme fixe exécuté dans un processus séparé.
# Le contenu fourni par l'IA est du HTML, jamais du code Python.
MOTEUR = r"""
import json
import resource
import subprocess
import sys

# Limites pour le rendu et la fabrication des aperçus.
resource.setrlimit(resource.RLIMIT_CPU, (30, 30))
resource.setrlimit(
    resource.RLIMIT_AS,
    (2 * 1024**3, 2 * 1024**3),
)
resource.setrlimit(
    resource.RLIMIT_FSIZE,
    (12 * 1024**2, 12 * 1024**2),
)
resource.setrlimit(resource.RLIMIT_CORE, (0, 0))

demande = json.load(sys.stdin)

if demande["action"] == "pdf":
    from weasyprint import HTML, CSS
    from weasyprint.urls import URLFetcher, FatalURLFetchingError

    class SansRessourcesExternes(URLFetcher):
        def fetch(self, url, headers=None):
            raise FatalURLFetchingError(
                "Ressource externe interdite. "
                "Utiliser du CSS et du SVG directement intégrés au HTML."
            )

    fetcher = SansRessourcesExternes()

    # Valeurs de secours seulement :
    # le CSS contenu dans le HTML peut les remplacer.
    base = CSS(
                string='''
            @page {
                size: A4;
                margin: 18mm;
            }

            body {
                font-family: "DejaVu Sans", sans-serif;
                font-size: 11pt;
                line-height: 1.35;
            }

            svg, img {
                max-width: 100%;
            }
                ''',
        url_fetcher=fetcher,
    )

    document = HTML(
        string=demande["html"],
        base_url="https://ressources.invalid/",
        url_fetcher=fetcher,
    ).render(stylesheets=[base])

    nombre_pages = len(document.pages)

    if not 1 <= nombre_pages <= 30:
        raise ValueError("Le document doit contenir entre 1 et 30 pages.")

    # Ne pas inclure de fichiers joints dans le PDF.
    document.metadata.attachments = []

    document.write_pdf("resultat.pdf")

    print(json.dumps({"pages": nombre_pages}))

elif demande["action"] == "apercu":
    subprocess.run(
        [
            "/usr/bin/pdftoppm",
            "-f", str(demande["page"]),
            "-l", str(demande["page"]),
            "-singlefile",
            "-scale-to", "1600",
            "-png",
            demande["fichier"],
            "apercu",
        ],
        check=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        timeout=25,
    )

    print("{}")

else:
    raise ValueError("Action inconnue.")
"""


def executer_moteur(demande: dict, sortie: str) -> tuple[bytes, dict]:
    if not VERROU_RENDU.acquire(blocking=False):
        raise ValueError(
            "Un rendu est déjà en cours. Réessayer après sa fin."
        )

    try:
        with TemporaryDirectory(prefix="documents-ia-") as temporaire:
            dossier = Path(temporaire)

            environnement = {
                "PATH": "/usr/bin:/bin",
                "HOME": str(dossier),
                "LANG": "C.UTF-8",
                "TMPDIR": str(dossier),
                "XDG_CACHE_HOME": str(dossier),
            }

            with (dossier / "journal.txt").open("w+b") as journal:
                try:
                    resultat = subprocess.run(
                        [sys.executable, "-I", "-c", MOTEUR],
                        input=json.dumps(
                            demande,
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        stdout=subprocess.PIPE,
                        stderr=journal,
                        cwd=dossier,
                        env=environnement,
                        timeout=45,
                    )
                except subprocess.TimeoutExpired:
                    raise ValueError(
                        "Rendu interrompu après 45 secondes. "
                        "Simplifier le document."
                    ) from None

                journal.seek(0, 2)
                journal.seek(max(0, journal.tell() - 1800))

                details = journal.read().decode(
                    "utf-8",
                    errors="replace",
                ).strip()

            if resultat.returncode != 0:
                raise ValueError(
                    "Échec du moteur de mise en page. "
                    + (details or "Limite de ressources atteinte.")
                )

            fichier_sortie = dossier / sortie

            if not fichier_sortie.is_file():
                raise ValueError(
                    "Le moteur n'a pas produit le fichier attendu."
                )

            informations = json.loads(
                resultat.stdout.decode("utf-8")
            )
            informations["avertissements_moteur"] = details

            return fichier_sortie.read_bytes(), informations

    finally:
        VERROU_RENDU.release()


def retrouver_pdf(fichier: str) -> Path:
    nom = fichier.removeprefix("documents_generes/")

    # Anciens fichiers comportant un identifiant aléatoire.
    ancien_format = (
        r"[A-Za-z0-9_-]{1,60}-[0-9a-f]{12}\.pdf"
    )

    # Nouveaux fichiers comportant une date et une heure.
    nouveau_format = (
        r"[A-Za-z0-9_-]{1,60}_"
        r"[0-9]{4}-[0-9]{2}-[0-9]{2}_"
        r"[0-9]{2}h[0-9]{2}m[0-9]{2}s"
        r"(?:-v[0-9]{2,})?\.pdf"
    )

    if not (
        re.fullmatch(ancien_format, nom)
        or re.fullmatch(nouveau_format, nom)
    ):
        raise ValueError(
            "Utiliser le nom exact d'un PDF créé par creer_pdf."
        )

    chemin = DOSSIER / nom

    if chemin.is_symlink() or not chemin.is_file():
        raise ValueError("PDF introuvable.")

    return chemin



@mcp.tool()
def creer_pdf(
    html: str,
    nom: str = "document",
    profil: Literal[
        "chatgpt", "claude", "gemini", "libre"
    ] = "libre",
) -> dict:
    """Crée un PDF à partir d'un document HTML complet.

    Le HTML doit contenir la mise en page et le CSS dans <style>.
    Les illustrations peuvent être du SVG directement intégré.

    Aucun chargement de ressources externes :
    pas de fichier local, image distante, police distante ou CDN.

    Le profil est une étiquette descriptive.
    Il ne remplace pas le travail graphique de l'IA :
    l'apparence dépend du HTML et du CSS fournis.

    Pour modifier un document existant :
    lire_source_pdf, modifier son HTML, puis créer une nouvelle version.

    Le résultat est un chemin dans Codespaces, pas un lien web.
    La création ne constitue pas une vérification visuelle.
    """
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,60}", nom):
        raise ValueError(
            "Nom : 1 à 60 lettres sans accents, chiffres, "
            "tirets ou underscores."
        )

    if not html.strip() or len(html.encode("utf-8")) > 800_000:
        raise ValueError(
            "Le HTML doit être non vide et inférieur à 800 Ko."
        )

    if len(list(DOSSIER.glob("*.pdf"))) >= 200:
        raise ValueError(
            "Limite de 200 PDF atteinte. "
            "Supprimer les anciennes versions inutiles."
        )

    contenu, informations = executer_moteur(
        {"action": "pdf", "html": html},
        "resultat.pdf",
    )

    creation = datetime.now(ZoneInfo("Europe/Paris"))
    base = f"{nom}_{creation:%Y-%m-%d_%Hh%Mm%Ss}"

    contenu_source = json.dumps(
        {
            "profil": profil,
            "pages": informations["pages"],
            "cree_le": creation.isoformat(timespec="seconds"),
            "fuseau_horaire": "Europe/Paris",
            "html": html,
        },
        ensure_ascii=False,
        indent=2,
    )

    version = 1

    while True:
        suffixe = "" if version == 1 else f"-v{version:02d}"
        identifiant = f"{base}{suffixe}"

        pdf = DOSSIER / f"{identifiant}.pdf"
        source = DOSSIER / f"{identifiant}.source.json"

        pdf_cree = False
        source_creee = False

        try:
            # Création exclusive : ne jamais écraser un fichier existant.
            with source.open("x", encoding="utf-8") as flux_source:
                source_creee = True
                source.chmod(0o600)

                with pdf.open("xb") as flux_pdf:
                    pdf_cree = True
                    pdf.chmod(0o600)

                    flux_pdf.write(contenu)
                    flux_source.write(contenu_source)

        except FileExistsError:
            # Ne supprimer que les fichiers créés par cette tentative.
            if pdf_cree:
                pdf.unlink(missing_ok=True)

            if source_creee:
                source.unlink(missing_ok=True)

            version += 1
            continue

        except Exception:
            # Nettoyage si l'écriture échoue.
            if pdf_cree:
                pdf.unlink(missing_ok=True)

            if source_creee:
                source.unlink(missing_ok=True)

            raise

        break


    return {
        "statut": "cree",
        "fichier": f"documents_generes/{pdf.name}",
        "profil": profil,
        "pages": informations["pages"],
        "taille_octets": pdf.stat().st_size,
        "avertissements_moteur": informations[
            "avertissements_moteur"
        ],
        "rendu_visuel_verifie": False,
        "recuperation": (
            "Télécharger le PDF depuis l'explorateur de Codespaces."
        ),
    }


@mcp.tool()
def lire_source_pdf(fichier: str) -> dict:
    """Relit le HTML d'un PDF créé par creer_pdf pour le modifier.

    Ne lit pas le contenu d'un PDF quelconque.
    Utiliser le chemin exact renvoyé lors de la création.
    """
    pdf = retrouver_pdf(fichier)
    source = pdf.with_suffix(".source.json")

    if (
        source.is_symlink()
        or not source.is_file()
        or source.stat().st_size > 6_000_000
    ):
        raise ValueError("Source du document indisponible.")

    return json.loads(source.read_text(encoding="utf-8"))


@mcp.tool()
def apercu_pdf(fichier: str, page: int = 1) -> Image:
    """Renvoie l'image PNG d'une page d'un PDF créé par creer_pdf.

    Sert au contrôle visuel de la composition.
    L'affichage et la lecture de cette image dépendent du client MCP.
    """
    pdf = retrouver_pdf(fichier)
    informations = lire_source_pdf(fichier)

    if not 1 <= page <= informations["pages"]:
        raise ValueError("Numéro de page inexistant.")

    image, _ = executer_moteur(
        {
            "action": "apercu",
            "fichier": str(pdf),
            "page": page,
        },
        "apercu.png",
    )

    return Image(data=image, format="png")


@mcp.tool()
def lister_pdf() -> list[str]:
    """Liste les PDF du dossier documents_generes."""
    return sorted(
        f"documents_generes/{f.name}"
        for f in DOSSIER.glob("*.pdf")
        if f.is_file() and not f.is_symlink()
    )
