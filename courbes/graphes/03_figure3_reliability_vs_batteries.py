import numpy as np
import matplotlib.pyplot as plt
from common import dss, build_default_params, save_figure


def main():
    params = build_default_params()
    lolp_cible = dss.TYPES_MENAGE[dss.TYPE_MENAGE_PAR_DEFAUT]["lolp_cible"]
    batteries_range = list(range(1, 15))
    fiabilites = []

    for n_batteries in batteries_range:
        resultats = dss.simuler_installation(28, n_batteries, 200, params, lolp_cible=lolp_cible)
        fiabilites.append(float(np.mean(resultats["fiabilite"])))
        print(f"{n_batteries} batteries -> fiabilité {fiabilites[-1]:.3f}")

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(batteries_range, fiabilites, marker="o", linestyle="-", color="#F58518")
    target_line = 1.0 - lolp_cible
    ax.axhline(target_line, color="#D9534F", linestyle="--", linewidth=1.5, label="Seuil cible 95%")
    ax.set_title("Figure 3 — Fiabilité simulée en fonction du nombre de batteries (panneaux fixés à 28)")
    ax.set_xlabel("Nombre de batteries")
    ax.set_ylabel("Fiabilité moyenne simulée")
    ax.set_ylim(0.0, 1.0)
    ax.grid(alpha=0.3)
    ax.legend()
    fig.tight_layout()
    save_figure(fig, "03_figure3_reliability_vs_batteries.png")


if __name__ == "__main__":
    main()
