import numpy as np
import matplotlib.pyplot as plt
from common import dss, build_default_params, save_figure


def modifier_matrice_psh(psh_matrice, prob_stable):
    matrice = np.array(psh_matrice, dtype=float)
    base_row = np.array(matrice[2], dtype=float)
    reste = base_row[:2]
    if reste.sum() <= 0:
        raise ValueError("La somme des probabilités de transition vers les états non pluvieux est nulle.")
    echelle = (1.0 - prob_stable) / reste.sum()
    matrice[2, 0] = reste[0] * echelle
    matrice[2, 1] = reste[1] * echelle
    matrice[2, 2] = prob_stable
    return matrice


def main():
    params = build_default_params()
    lolp_cible = dss.TYPES_MENAGE[dss.TYPE_MENAGE_PAR_DEFAUT]["lolp_cible"]
    categories = [
        ("Optimiste", 0.30),
        ("Nominal", 0.40),
        ("Pessimiste", 0.55),
        ("Dégradé (+30% nuageux)", 0.70),
    ]
    fiabilites = []

    for label, p_stable in categories:
        params_loc = dict(params)
        params_loc["psh_matrice"] = modifier_matrice_psh(params["psh_matrice"], p_stable).tolist()
        resultats = dss.simuler_installation(28, 4, 200, params_loc, lolp_cible=lolp_cible)
        fi = float(np.mean(resultats["fiabilite"]))
        fiabilites.append(fi)
        print(f"{label} ({p_stable:.2f}) -> fiabilité {fi:.3f}")

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(
        [label for label, _ in categories],
        [100.0 * v for v in fiabilites],
        color=["#6A7BE2", "#1F9977", "#6A7BE2", "#D96A2B"],
        edgecolor="black",
        width=0.65,
    )
    target_pct = 100.0 * (1.0 - lolp_cible)
    ax.axhline(target_pct, color="#D9534F", linestyle="--", linewidth=1.5, label="Seuil cible 95%")
    ax.set_title("Analyse de sensibilité — installation traditionnelle (28 panneaux / 4 batteries) selon la sévérité de la saison des pluies")
    ax.set_ylabel("Fiabilité simulée (%)")
    ax.set_ylim(0, 100)
    ax.set_yticks(np.linspace(0, 100, 6))
    ax.grid(axis="y", alpha=0.25)

    for bar, val in zip(bars, fiabilites):
        ax.text(bar.get_x() + bar.get_width() / 2, 100.0 * val + 2, f"{100.0 * val:.0f}%", ha="center", va="bottom", fontsize=10)

    ax.legend(loc="upper right")
    fig.tight_layout()
    save_figure(fig, "04_figure4_sensitivity_rainy_season_severity.png")


if __name__ == "__main__":
    main()
