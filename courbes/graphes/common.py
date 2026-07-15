import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

try:
    from modules import dualsolarstat_simulation as dss
except ImportError as exc:
    raise ImportError(
        "Impossible d'importer modules.dualsolarstat_simulation. "
        "Exécutez ce script depuis la racine du dépôt ou ajoutez le répertoire racine à PYTHONPATH."
    ) from exc

import numpy as np
import matplotlib.pyplot as plt

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def save_figure(fig, filename):
    chemin = os.path.join(OUTPUT_DIR, filename)
    fig.savefig(chemin, dpi=200, bbox_inches="tight")
    print(f"Saved {chemin}")
    plt.close(fig)
    return chemin


def build_default_params():
    parametres, _ = dss.construire_parametres(fichier_consommation=None, fichier_meteo=None)
    return parametres
