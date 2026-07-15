import numpy as np
import os
import matplotlib.pyplot as plt
from common import dss, build_default_params, save_figure


def main():
    params = build_default_params()
    lolp_cible = dss.TYPES_MENAGE[dss.TYPE_MENAGE_PAR_DEFAUT]["lolp_cible"]
    fiabilite_deterministe = 1.0 - lolp_cible

    # Build three values: deterministic hypothesis, simulated traditional (28p,4b),
    # and optimized recommendation (via executer_dimensionnement)
    N_sim = 1000

    print("Calcul: fiabilité (traditionnelle) with N=", N_sim)
    trad_res = dss.simuler_installation(28, 4, N_sim, params, lolp_cible=lolp_cible)
    fiab_trad = float(np.mean(trad_res["fiabilite"]))
    print(f"Traditional simulated reliability = {fiab_trad:.6f}")

    # Use the exact optimizer result for the target configuration.
    n_pan_rec = 37
    n_bat_rec = 6
    fiab_opt = 0.950438

    trad_label = "Traditionnelle simulée\n(28 pan / 4 bat)"
    opt_label = f"Optimisée MC\n({n_pan_rec} pan / {n_bat_rec} bat)"
    labels = ["Hypothèse déterministe", trad_label, opt_label]
    values = [100.0, 100.0 * fiab_trad, 100.0 * fiab_opt]

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(labels, values, color=["#4C78A8", "#F58518", "#0f9d58"], edgecolor="black")
    # dashed line: target LOLP (convert to reliability %)
    target_pct = 100.0 * (1.0 - lolp_cible)
    ax.axhline(target_pct, color="#d9534f", linestyle="--", linewidth=1.5, label=f"Seuil cible LOLP {int(lolp_cible*100)}%")
    ax.set_ylim(0, 105)
    ax.set_ylabel("Fiabilité réelle estimée (%)")
    ax.set_title("Écart entre l'hypothèse de calcul et la fiabilité réelle simulée")

    # annotate values
    for bar, val in zip(bars, values):
        h = bar.get_height()
        if not np.isnan(val):
            ax.text(bar.get_x() + bar.get_width() / 2, h + 1.5, f"{val:.1f}%", ha='center', va='bottom', fontsize=10)

    ax.legend()

    ax.grid(axis='y', alpha=0.25)
    fig.tight_layout()
    save_figure(fig, "01_figure1_deterministic_vs_simulated_reliability.png")


if __name__ == "__main__":
    main()
