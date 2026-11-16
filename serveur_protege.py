import hmac
import secrets
from pathlib import Path

import uvicorn
from starlette.responses import PlainTextResponse

from serveur_mcp import mcp
import outils_apercus
import outils_pptx
import outils_pdf
import controle_pdf
import telechargements
from telechargements import avec_telechargements
import outils_documents
import outils_excel


# La clé est stockée hors du dossier du projet.
dossier_cle = Path.home() / ".config" / "codespace-mcp"
dossier_cle.mkdir(parents=True, exist_ok=True)
dossier_cle.chmod(0o700)

fichier_cle = dossier_cle / "cle"
if not fichier_cle.exists():
    fichier_cle.write_text(secrets.token_urlsafe(48), encoding="utf-8")
    fichier_cle.chmod(0o600)

cle = fichier_cle.read_text(encoding="utf-8").strip()
if len(cle) < 40:
    raise RuntimeError("La clé MCP est absente ou trop courte.")


class ProtectionParCle:
    def __init__(self, application):
        self.application = application

    async def __call__(self, scope, receive, send):
        if scope["type"] == "http":
            autorisations = [
                valeur
                for nom, valeur in scope.get("headers", [])
                if nom.lower() == b"authorization"
            ]

            attendu = f"Bearer {cle}".encode("utf-8")

            if (
                len(autorisations) != 1
                or not hmac.compare_digest(autorisations[0], attendu)
            ):
                reponse = PlainTextResponse(
                    "Accès refusé",
                    status_code=401
                )
                await reponse(scope, receive, send)
                return

        await self.application(scope, receive, send)


application = ProtectionParCle(mcp.streamable_http_app())

if __name__ == "__main__":
    uvicorn.run(avec_telechargements(application),  # type: ignore
        host="127.0.0.1",
        port=8000,
        log_level="info",
    )
