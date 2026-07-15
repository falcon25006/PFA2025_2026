import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.cm import viridis
from matplotlib.colors import LogNorm

# ============================================================
# À MODIFIER : chemin vers ton fichier CSV
# ============================================================
CSV_PATH = os.path.join(os.path.dirname(__file__), "output", "06_robustesse_boxplot_data.csv")

output_dir = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(output_dir, exist_ok=True)

if not os.path.exists(CSV_PATH):
    raise FileNotFoundError(f"CSV non trouvé: {CSV_PATH}")

df = pd.read_csv(CSV_PATH)

# Compter les occurrences de chaque combinaison (n_batteries, n_panneaux)
counts = df.groupby(['n_batteries', 'n_panneaux']).size().reset_index(name='count')
counts = counts.sort_values('count', ascending=False).reset_index(drop=True)

# Etiquette lisible par config
labels = [f"{b} bat. / {p} pan." for b, p in zip(counts['n_batteries'], counts['n_panneaux'])]

# Couleur des barres selon le compte (même échelle que la version scatter)
norm = LogNorm(vmin=1, vmax=counts['count'].max())
colors = viridis(norm(counts['count']))

fig, ax = plt.subplots(figsize=(7.5, 5), dpi=200)

bars = ax.bar(labels, counts['count'], color=colors, edgecolor='white', linewidth=0.6)

# Etiquettes de valeur au-dessus de chaque barre
for rect, val in zip(bars, counts['count']):
    ax.text(rect.get_x() + rect.get_width() / 2, rect.get_height(),
            str(val), ha='center', va='bottom', fontsize=9)

ax.set_yscale('log')
ax.set_xlabel('Configuration (batteries / panneaux)')
ax.set_ylabel("Nombre d'exécutions (échelle log)")
ax.set_title("Configurations trouvées par la recherche jointe\n(1000 exécutions)")

plt.xticks(rotation=30, ha='right')
ax.grid(True, axis='y', alpha=0.3)
fig.tight_layout()

output_path = os.path.join(output_dir, "06_figure6_robustesse_boxplot.png")
fig.savefig(output_path, dpi=200)
print(f"✓ Image sauvegardée: {output_path}")
plt.close()