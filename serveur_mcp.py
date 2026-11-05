import os

from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings

nom_codespace = os.environ.get("CODESPACE_NAME", "")
domaine = os.environ.get(
    "GITHUB_CODESPACES_PORT_FORWARDING_DOMAIN",
    "app.github.dev"
)

if not nom_codespace:
    raise RuntimeError("Le nom du Codespace est introuvable.")

hote_codespace = f"{nom_codespace}-8000.{domaine}"

mcp = FastMCP(
    "Codespace Documents",
    host="127.0.0.1",
    port=8000,
    stateless_http=True,
    json_response=True,
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=[
            hote_codespace,
            "127.0.0.1:8000",
            "localhost:8000",
        ],
        allowed_origins=[],
    ),
)

@mcp.tool()
def verifier_connexion() -> str:
    """Vérifie la connexion au serveur MCP du Codespace."""
    return "Connexion réussie au Codespace !"

if __name__ == "__main__":
    mcp.run(transport="streamable-http")
