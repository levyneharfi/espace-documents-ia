"""Images en mémoire : PNG/JPEG base64 ou SVG autonome, sans URL externe."""
import base64
import binascii
import io
import subprocess
import sys
import warnings
from PIL import Image

LIMITE = 4_000_000
Image.MAX_IMAGE_PIXELS = 16_000_000

def image_png(image_base64='', svg=''):
    if bool(image_base64) == bool(svg):
        raise ValueError('Fournir exactement image_base64 OU svg.')
    if svg:
        if len(svg.encode('utf-8')) > 300_000:
            raise ValueError('SVG limité à 300 Ko.')
        # Conversion isolée, limites CPU/mémoire et délai maximal.
        programme = r"""
import resource, sys, re
resource.setrlimit(resource.RLIMIT_CPU, (10, 10))
resource.setrlimit(resource.RLIMIT_AS, (1024**3, 1024**3))
from defusedxml import ElementTree as ET
from cairosvg.parser import Tree
from cairosvg.surface import PNGSurface
contenu = sys.stdin.buffer.read(300001)
root = ET.fromstring(contenu)
if root.tag != '{http://www.w3.org/2000/svg}svg':
    raise ValueError('Racine SVG avec namespace requise.')
for element in root.iter():
    for nom, valeur in element.attrib.items():
        if nom.split('}')[-1].lower() in ('href', 'src') and not valeur.strip().startswith('#'):
            raise ValueError('Références externes interdites.')
def refuse(*args, **kwargs):
    raise ValueError('Chargement de ressource interdit.')
arbre = Tree(bytestring=contenu, url_fetcher=refuse, unsafe=False)
surface = PNGSurface(arbre, None, 96, output_width=1600, output_height=1600)
surface.finish()
sys.stdout.buffer.write(surface.cairo.getvalue())
"""
        # Utiliser BytesIO explicitement plutôt que dépendre d'un attribut interne.
        programme = programme.replace(
            'surface = PNGSurface(arbre, None, 96, output_width=1600, output_height=1600)\nsurface.finish()\nsys.stdout.buffer.write(surface.cairo.getvalue())',
            'import io\nsortie = io.BytesIO()\nsurface = PNGSurface(arbre, sortie, 96, output_width=1600)\nsurface.finish()\nsys.stdout.buffer.write(sortie.getvalue())'
        )
        resultat = subprocess.run(
            [sys.executable, '-I', '-c', programme],
            input=svg.encode('utf-8'), capture_output=True, timeout=15,
        )
        if resultat.returncode:
            raise ValueError('Conversion SVG échouée : SVG autonome requis.')
        donnees = resultat.stdout
    else:
        if len(image_base64) > 5_400_000:
            raise ValueError('Image base64 trop volumineuse.')
        try:
            donnees = base64.b64decode(image_base64, validate=True)
        except (ValueError, binascii.Error):
            raise ValueError('Base64 brut invalide, sans préfixe data:.') from None
    if len(donnees) > LIMITE:
        raise ValueError('Image limitée à 4 Mo.')
    with warnings.catch_warnings():
        warnings.simplefilter('error', Image.DecompressionBombWarning)
        with Image.open(io.BytesIO(donnees)) as im:
            if im.format not in ('PNG', 'JPEG'):
                raise ValueError('Formats acceptés : PNG et JPEG.')
            if im.width * im.height > 16_000_000:
                raise ValueError('Image limitée à 16 millions de pixels.')
            im.load()
            im = im.convert('RGBA')
            sortie = io.BytesIO()
            im.save(sortie, format='PNG')
            return sortie.getvalue()
