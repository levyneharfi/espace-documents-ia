"""Liens de téléchargement signés et temporaires pour les fichiers générés."""
import hashlib
import hmac
import os
import secrets
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
from zoneinfo import ZoneInfo

from starlette.responses import FileResponse, PlainTextResponse

from serveur_mcp import mcp

DOSSIER = (Path(__file__).resolve().parent / "documents_generes").resolve()
FICHIER_SECRET = Path.home() / ".config" / "codespace-mcp" / "secret_liens"
TYPES = {
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}
PREFIXE = "/telecharger/"
DUREE_MAX_HEURES = 72


def _secret() -> bytes:
    if not FICHIER_SECRET.exists():
        FICHIER_SECRET.parent.mkdir(parents=True, exist_ok=True)
        fd = os.open(FICHIER_SECRET, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "w") as f:
            f.write(secrets.token_hex(32))
    return bytes.fromhex(FICHIER_SECRET.read_text().strip())


def _signature(nom: str, expiration: int) -> str:
    message = f"{nom}\n{expiration}".encode()
    return hmac.new(_secret(), message, hashlib.sha256).hexdigest()


def _fichier_autorise(nom: str):
    if not nom or "/" in nom or "\\" in nom or nom.startswith("."):
        return None
    chemin = (DOSSIER / nom).resolve()
    if chemin.parent != DOSSIER or chemin.suffix.lower() not in TYPES:
        return None
    if not chemin.is_file():
        return None
    return chemin


def _url_publique() -> str:
    nom = os.environ.get("CODESPACE_NAME")
    domaine = os.environ.get("GITHUB_CODESPACES_PORT_FORWARDING_DOMAIN", "app.github.dev")
    if not nom:
        raise RuntimeError("CODESPACE_NAME introuvable : le serveur ne tourne pas dans un Codespace.")
    return f"https://{nom}-8000.{domaine}"


def creer_lien(fichier: str, duree_heures: int = 24) -> dict:
    nom = Path(fichier).name
    chemin = _fichier_autorise(nom)
    if chemin is None:
        raise ValueError(f"Fichier introuvable ou non autorisé dans documents_generes : {nom}")
    duree = max(1, min(int(duree_heures), DUREE_MAX_HEURES))
    expiration = int(time.time()) + duree * 3600
    url = f"{_url_publique()}{PREFIXE}{expiration}/{_signature(nom, expiration)}/{quote(nom)}"
    expire = datetime.fromtimestamp(expiration, ZoneInfo("Europe/Paris"))
    return {
        "fichier": nom,
        "url": url,
        "expire_le": expire.strftime("%d/%m/%Y à %Hh%M (heure de Paris)"),
        "taille_octets": chemin.stat().st_size,
        "rappel": "Le lien fonctionne uniquement si le Codespace et le serveur sont allumés.",
    }


@mcp.tool()
def lien_telechargement(fichier: str, duree_heures: int = 24) -> dict:
    """Crée un lien de téléchargement direct et temporaire pour un fichier de documents_generes
    (pdf, docx, xlsx, pptx). Donner le nom ou le chemin exact renvoyé par l'outil de création.
    duree_heures : 1 à 72 (24 par défaut). Recopier l'URL renvoyée exactement, sans la modifier."""
    return creer_lien(fichier, duree_heures)


def _reponse(scope):
    if scope.get("method") not in ("GET", "HEAD"):
        return PlainTextResponse("Méthode non autorisée.", status_code=405)
    morceaux = scope["path"][len(PREFIXE):].split("/")
    refus = PlainTextResponse("Lien invalide.", status_code=403)
    if len(morceaux) != 3 or not morceaux[0].isdigit():
        return refus
    expiration_txt, signature, nom = morceaux
    expiration = int(expiration_txt)
    if not hmac.compare_digest(signature, _signature(nom, expiration)):
        return refus
    if time.time() > expiration:
        return PlainTextResponse("Ce lien a expiré : demande un nouveau lien à l'IA.", status_code=410)
    chemin = _fichier_autorise(nom)
    if chemin is None:
        return PlainTextResponse("Fichier introuvable.", status_code=404)
    return FileResponse(
        chemin,
        media_type=TYPES[chemin.suffix.lower()],
        filename=nom,
        headers={"Cache-Control": "no-store", "X-Content-Type-Options": "nosniff"},
    )


def avec_telechargements(app_suivante):
    """Traite /telecharger/... avant la protection par clé et transmet tout le reste sans changement."""
    async def app(scope, receive, send):
        if scope["type"] == "http" and scope["path"].startswith(PREFIXE):
            await _reponse(scope)(scope, receive, send)
            return
        await app_suivante(scope, receive, send)
    return app
