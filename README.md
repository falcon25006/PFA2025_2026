# PFA 2025-2026 - DualSolarStat

Application de dimensionnement solaire probabiliste basée sur une interface graphique CustomTkinter et un moteur de calcul Monte-Carlo.

## Présentation

Ce projet compare le dimensionnement solaire traditionnel et probabiliste, en intégrant :

- lecture de données CSV pour la consommation et l'ensoleillement
- charge d'un catalogue matériel (panneaux, batteries, onduleurs)
- simulation Monte-Carlo et analyse de convergence
- export de résultats au format CSV et DOCX
- interface graphique complète avec onglets de simulation

## Structure du projet

- `main.py` : point d'entrée principal qui lance le module `modules.dualsolarstat_simulation`
- `modules/dualsolarstat_simulation.py` : application principale avec calculs, GUI et export
- `requirements.txt` : dépendances Python
- `courbes/` : scripts de génération de figures et d'analyses
- `files/` : espace pour les données ou ressources supplémentaires
- `temp_catalogue_test.csv` : exemple de catalogue matériel
- `temp_doc_inspect.py` : script utilitaire temporaire

## Installation

1. Créer et activer un environnement virtuel :

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Installer les dépendances :

```powershell
pip install -r requirements.txt
```

## Exécution

Lancer l'application depuis la racine du projet :

```powershell
python main.py
```

Cela ouvre l'interface graphique `DualSolarStat`.

## Dépendances

- `customtkinter`
- `matplotlib`
- `numpy`
- `pandas`
- `python-docx`

> Si `python-docx` n'est pas dans `requirements.txt`, installez-le manuellement avec `pip install python-docx`.

## Notes

- Les fichiers de données doivent être des CSV encodés en UTF-8 ou UTF-8-BOM.
- Le catalogue matériel doit contenir les colonnes `type`, `nom`, `caracteristique`, `prix`.
- Le module principal est conçu pour être lancé via `main.py`.
