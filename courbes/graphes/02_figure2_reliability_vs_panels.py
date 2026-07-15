import numpy as np
import matplotlib.pyplot as plt
from common import dss, build_default_params, save_figure


def main():
    params = build_default_params()
    lolp_cible = dss.TYPES_MENAGE[dss.TYPE_MENAGE_PAR_DEFAUT]["lolp_cible"]
    panneaux_range = list(range(16, 45, 2))
    fiabilites = []

    for n_panneaux in panneaux_range:
        resultats = dss.simuler_installation(n_panneaux, 4, 200, params, lolp_cible=lolp_cible)
        fiabilites.append(float(np.mean(resultats["fiabilite"])))
        print(f"{n_panneaux} panneaux -> fiabilité {fiabilites[-1]:.3f}")

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(panneaux_range, fiabilites, marker="o", linestyle="-", color="#4C78A8")
    target_line = 1.0 - lolp_cible
    ax.axhline(target_line, color="#D9534F", linestyle="--", linewidth=1.5, label="Seuil cible 95%")
    ax.set_title("Figure 2 — Fiabilité simulée en fonction du nombre de panneaux (batteries fixées à 4)")
    ax.set_xlabel("Nombre de panneaux")
    ax.set_ylabel("Fiabilité moyenne simulée")
    ax.set_ylim(0.0, 1.0)
    ax.grid(alpha=0.3)
    ax.legend()
    fig.tight_layout()
    save_figure(fig, "02_figure2_reliability_vs_panels.png")


if __name__ == "__main__":
    main()
