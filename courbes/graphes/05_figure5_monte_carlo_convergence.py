import numpy as np
import matplotlib.pyplot as plt
from common import dss, build_default_params, save_figure


def main():
    params = build_default_params()
    lolp_cible = dss.TYPES_MENAGE[dss.TYPE_MENAGE_PAR_DEFAUT]["lolp_cible"]
    tailles = [50, 100, 200, 400, 800, 1600, 3200, 6400]
    # Number of repeated estimator evaluations per S to compute an empirical IC95%
    repeats_per_s = 30
    means = []
    sems = []

    for s in tailles:
        vals = []
        for i in range(repeats_per_s):
            resultats = dss.simuler_installation(28, 4, s, params, lolp_cible=lolp_cible)
            vals.append(float(np.mean(resultats["fiabilite"])))
        arr = np.array(vals)
        mean = float(arr.mean())
        sem = float(arr.std(ddof=1) / np.sqrt(len(arr)))
        means.append(mean)
        sems.append(sem)
        print(f"S={s} -> mean={mean:.6f}, sem={sem:.6f} (n={len(arr)})")

    # Reference value computed with a large N
    ref_res = dss.simuler_installation(28, 4, 6000, params, lolp_cible=lolp_cible)
    ref_val = float(np.mean(ref_res["fiabilite"]))
    print(f"Référence (N=6000) = {ref_val:.6f}")

    # convert to percentages for plotting like the Word figure
    tailles_arr = np.array(tailles)
    means_pct = 100.0 * np.array(means)
    sems_pct = 100.0 * np.array(sems)
    target_pct = 100.0 * (1.0 - lolp_cible)

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.errorbar(tailles_arr, means_pct, yerr=1.96 * sems_pct, fmt='-o', color='#0f9d58', capsize=5, label='Estimation ± IC 95%')
    ax.set_xscale('log')
    ax.axhline(target_pct, color='#D9534F', linestyle='--', linewidth=1.5, label="Seuil cible 95%")
    ax.set_title("Convergence de l'estimateur Monte-Carlo (installation traditionnelle)")
    ax.set_xlabel("Nombre de simulations Monte-Carlo (N)")
    ax.set_ylabel("Fiabilité estimée (%)")
    ax.set_ylim(np.min(means_pct - 1.96 * sems_pct) - 0.5, np.max(means_pct + 1.96 * sems_pct) + 0.5)
    ax.grid(True, which='both', alpha=0.3)
    ax.legend()
    fig.tight_layout()
    save_figure(fig, "05_figure5_monte_carlo_convergence.png")


if __name__ == "__main__":
    main()
