"""Aperçus de documents générés : PDF, DOCX, XLSX et PPTX.

Réservé aux documents produits par le serveur, pas aux fichiers externes inconnus.
LibreOffice n'est pas une sandbox. Ne pas exposer la conversion de fichiers
arbitraires sans isolation réseau/processus supplémentaire.
"""
import hashlib
import json
import os
import re
import shutil
import signal
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Lock
from typing import Literal

from mcp.server.fastmcp import Image
from serveur_mcp import mcp

RACINE = Path(__file__).resolve().parent
DOSSIER = (RACINE / 'documents_generes').resolve()
CACHE = DOSSIER / 'apercus'
CACHE.mkdir(parents=True, exist_ok=True)
VERROU = Lock()
EXTENSIONS = {'.pdf', '.docx', '.xlsx', '.pptx'}
MAX_PAGES = 120


def _trouver(fichier: str) -> Path:
    from rangement_documents import trouver_document
    chemin = trouver_document(fichier, EXTENSIONS)
    if chemin.stat().st_size > 30_000_000:
        raise ValueError('Document limité à 30 Mo pour les aperçus.')
    return chemin


def _executer(args: list[str], secondes: int, cwd: Path | None = None) -> str:
    # Groupe de processus dédié : arrêter aussi les enfants lors d'un dépassement.
    with TemporaryDirectory(prefix='journal-apercu-') as tmp:
        log = Path(tmp) / 'journal.txt'
        with log.open('w+b') as flux:
            environnement = os.environ.copy()
            environnement['LC_ALL'] = 'C'
            environnement['SAL_USE_VCLPLUGIN'] = 'svp'
            p = subprocess.Popen(
                args, cwd=cwd, stdin=subprocess.DEVNULL,
                stdout=flux, stderr=subprocess.STDOUT,
                start_new_session=True, env=environnement,
            )
            try:
                p.wait(timeout=secondes)
            except subprocess.TimeoutExpired:
                try:
                    os.killpg(p.pid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
                p.wait()
                raise ValueError(f'Conversion interrompue après {secondes} secondes.') from None
            flux.seek(0, 2)
            taille = flux.tell()
            flux.seek(max(0, taille - 5000))
            sortie = flux.read().decode('utf-8', errors='replace')
        if p.returncode != 0:
            raise ValueError('Échec de conversion : ' + sortie[-1200:])
        return sortie


def _pages(pdf: Path) -> int:
    informations = _executer(['pdfinfo', str(pdf)], 20)
    match = re.search(r'^Pages:\s+(\d+)\s*$', informations, re.M)
    if not match:
        raise ValueError('Nombre de pages du PDF non reconnu.')
    nombre = int(match.group(1))
    if not 1 <= nombre <= MAX_PAGES:
        raise ValueError(
            f'Le rendu contient {nombre} pages ; limite : {MAX_PAGES}. '
            'Pour XLSX, revoir les zones et paramètres d’impression avant de réessayer.'
        )
    return nombre


def _preparer(fichier: str) -> tuple[Path, Path, int, bool]:
    origine = _trouver(fichier)
    empreinte = hashlib.sha256()
    empreinte.update(origine.name.encode('utf-8'))
    with origine.open('rb') as flux:
        for bloc in iter(lambda: flux.read(1024 * 1024), b''):
            empreinte.update(bloc)
    cle = empreinte.hexdigest()
    dossier = CACHE / origine.suffix.lstrip('.') / cle
    dossier.mkdir(parents=True, exist_ok=True)
    pdf = dossier / 'rendu.pdf'
    meta = dossier / 'informations.json'
    reutilise = pdf.is_file() and meta.is_file()
    if reutilise:
        try:
            nombre = json.loads(meta.read_text(encoding='utf-8'))['pages']
            if not isinstance(nombre, int) or not 1 <= nombre <= MAX_PAGES:
                raise ValueError('Cache invalide.')
            return origine, pdf, nombre, True
        except (ValueError, KeyError, TypeError):
            reutilise = False
    with TemporaryDirectory(prefix='conversion-doc-') as temporaire:
        travail = Path(temporaire)
        copie = travail / ('document' + origine.suffix)
        shutil.copyfile(origine, copie)
        if origine.suffix == '.pdf':
            produit = copie
        else:
            profil = (travail / 'profil-libreoffice').as_uri()
            sortie = travail / 'sortie'
            sortie.mkdir()
            filtres = {
                '.docx': 'pdf:writer_pdf_Export',
                '.xlsx': 'pdf:calc_pdf_Export',
                '.pptx': 'pdf:impress_pdf_Export',
            }
            _executer([
                'libreoffice', '-env:UserInstallation=' + profil,
                '--headless', '--nologo', '--nodefault', '--norestore',
                '--convert-to', filtres[origine.suffix],
                '--outdir', str(sortie), str(copie),
            ], 90, travail)
            produit = sortie / 'document.pdf'
            if not produit.is_file():
                raise ValueError('LibreOffice n’a pas produit le PDF attendu.')
        if produit.stat().st_size > 40_000_000:
            raise ValueError('PDF d’aperçu limité à 40 Mo.')
        nombre = _pages(produit)
        provisoire = dossier / 'rendu.tmp'
        shutil.copyfile(produit, provisoire)
        provisoire.chmod(0o600)
        provisoire.replace(pdf)
        meta.write_text(json.dumps({'pages': nombre}), encoding='utf-8')
        meta.chmod(0o600)
    return origine, pdf, nombre, reutilise


@mcp.tool()
def preparer_apercus(fichier: str) -> dict:
    """Prépare les aperçus d'un PDF, DOCX, XLSX ou PPTX créé par le serveur.

    Retourne le nombre de pages du rendu. Ne renvoie pas encore d'image :
    appeler ensuite apercu_document pour une page. Conversion mise en cache.
    DOCX/XLSX/PPTX : rendu LibreOffice, pas Microsoft Office.
    XLSX : pages imprimées de toutes les feuilles visibles, selon les paramètres
    d’impression enregistrés ; ce ne sont pas des captures de l'interface Excel.
    Aucun contrôle visuel n'est effectué par cet outil.
    """
    if not VERROU.acquire(blocking=False):
        raise ValueError('Un aperçu est déjà en préparation. Réessayer après sa fin.')
    try:
        origine, pdf, nombre, reutilise = _preparer(fichier)
        return {
            'fichier': origine.relative_to(RACINE).as_posix(),
            'pages_du_rendu': nombre,
            'moteur': 'PDF original' if origine.suffix == '.pdf' else 'LibreOffice',
            'cache_reutilise': reutilise,
            'rendu_visuel_verifie': False,
            'limite': 'Les différences de polices et de moteur peuvent modifier le rendu bureautique.',
            'etape_suivante': 'Appeler apercu_document sur la page 1 et vérifier que l’image est accessible avant de demander les autres pages.',
        }
    finally:
        VERROU.release()


@mcp.tool()
def apercu_document(
    fichier: str, page: int = 1,
    resolution: Literal[1000, 1600] = 1000,
) -> Image:
    """Renvoie en PNG une page d'un PDF, DOCX, XLSX ou PPTX généré.

    Préparation automatique et cache inclus. Pages numérotées à partir de 1.
    Pour DOCX/XLSX/PPTX, il s'agit du rendu LibreOffice.
    Tester d'abord une page. Si l'image n'est pas accessible au modèle,
    ne pas demander les autres et ne pas prétendre à un contrôle visuel.
    """
    if not VERROU.acquire(blocking=False):
        raise ValueError('Un aperçu est déjà en préparation. Réessayer après sa fin.')
    try:
        origine, pdf, nombre, reutilise = _preparer(fichier)
        if not 1 <= page <= nombre:
            raise ValueError(f'Page attendue : 1 à {nombre}.')
        png = pdf.parent / f'page-{page}-{resolution}.png'
        if not png.is_file():
            with TemporaryDirectory(prefix='page-doc-') as tmp:
                prefixe = Path(tmp) / 'page'
                _executer([
                    'pdftoppm', '-f', str(page), '-l', str(page),
                    '-singlefile', '-scale-to', str(resolution),
                    '-png', str(pdf), str(prefixe),
                ], 30)
                produit = prefixe.with_suffix('.png')
                if not produit.is_file() or produit.stat().st_size > 12_000_000:
                    raise ValueError('Aperçu absent ou trop volumineux.')
                shutil.copyfile(produit, png)
                png.chmod(0o600)
        return Image(data=png.read_bytes(), format='png')
    finally:
        VERROU.release()
