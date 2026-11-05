#!/usr/bin/env bash
cd /workspaces/espace-documents-ia || exit 1

rendre_public() {
  local ok=0 etat=""
  for essai in $(seq 1 20); do
    sleep 6
    gh codespace ports visibility 8000:public -c "$CODESPACE_NAME" >/dev/null 2>&1
    etat=$(gh codespace ports -c "$CODESPACE_NAME" --json sourcePort,visibility \
      -q '.[] | select(.sourcePort == 8000) | .visibility' 2>/dev/null)
    if [ "$etat" = "public" ]; then
      ok=$((ok + 1))
      if [ "$ok" -ge 2 ]; then
        echo ">>> Port 8000 : public (vérifié)"
        return 0
      fi
    else
      ok=0
    fi
  done
  echo ">>> ATTENTION : le port 8000 n'est pas public (état : ${etat:-introuvable})."
  echo ">>> Rends-le Public dans l'onglet Ports, ou tape : port-public"
}

if curl -s -o /dev/null --max-time 2 http://127.0.0.1:8000/mcp; then
  echo "Le serveur tourne déjà."
  rendre_public
  exit 0
fi
source .venv/bin/activate
rendre_public &
exec python serveur_protege.py
