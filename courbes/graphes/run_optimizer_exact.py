import os
from common import dss, build_default_params

def main():
    params = build_default_params()
    chemin_catalogue = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'temp_catalogue_test.csv'))
    config = {
        "fichier_catalogue": chemin_catalogue,
        "budget_max": 100_000_000,
        "surface_max_m2": 400.0,
        "type_menage": dss.TYPE_MENAGE_PAR_DEFAUT,
        "fichier_consommation": None,
        "fichier_meteo": None,
        # repetitions per tested config (higher -> more accurate, slower)
        "n_repetitions_par_config": 20,
    }
    print('Running executer_dimensionnement with config:', config)
    res = dss.executer_dimensionnement(config)
    print('Result:', res)

if __name__ == '__main__':
    main()
