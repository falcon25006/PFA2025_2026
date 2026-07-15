from common import dss, build_default_params


def main():
    params = build_default_params()
    lolp_cible = dss.TYPES_MENAGE[dss.TYPE_MENAGE_PAR_DEFAUT]["lolp_cible"]
    n_pan = 37
    n_bat = 6
    N_sim = 1000
    print(f"Simulating {n_pan} panneaux / {n_bat} batteries with N={N_sim}")
    resultats = dss.simuler_installation(n_pan, n_bat, N_sim, params, lolp_cible=lolp_cible)
    fiab = float(resultats["fiabilite"].mean())
    print(f"Fiabilité estimée = {fiab:.6f} ({fiab*100:.1f}%)")


if __name__ == '__main__':
    main()
