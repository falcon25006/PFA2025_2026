import csv
import os
from docx import Document
from modules.dualsolarstat_simulation import executer_dimensionnement

# Read first paragraphs from the Word document
path_doc = r'C:\Users\zougo\Desktop\IGIT_2025_2026\PFA2025_2026\files\Comparaison_methode_traditionnelle_vs_DualSolarStat.docx'
doc = Document(path_doc)
with open('temp_word_preview.txt', 'w', encoding='utf-8') as out:
    for i, p in enumerate(doc.paragraphs[:160], 1):
        out.write(f'{i}: {p.text}\n')

# Build a sample catalogue and run the dimensionnement
catalogue_path = os.path.abspath('temp_catalogue_test.csv')
with open(catalogue_path, 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f)
    w.writerow(['type', 'nom', 'caracteristique', 'prix'])
    w.writerow(['panneau', 'Panneau monocristallin 230Wc', 230, 77000])
    w.writerow(['panneau', 'Panneau monocristallin 400Wc', 400, 220000])
    w.writerow(['batterie', 'Batterie Lithium LiFePO4 51.2V 100Ah', 100, 570000])
    w.writerow(['batterie', 'Batterie Lithium LiFePO4 51.2V 200Ah', 200, 1800000])
    w.writerow(['batterie', 'Batterie Lithium LiFePO4 51.2V 400Ah', 400, 5400000])
    w.writerow(['onduleur', 'Onduleur hybride 7kW MPPT intégré', 7000, 945000])
    w.writerow(['onduleur', 'Onduleur hybride 5kW MPPT intégré', 5000, 820000])
    w.writerow(['onduleur', 'Onduleur hybride 3kW MPPT intégré', 3000, 620000])

config = {
    'fichier_catalogue': catalogue_path,
    'budget_max': 10000000,
    'type_menage': 'moyenne',
    'fichier_consommation': None,
    'fichier_meteo': None,
    'n_repetitions_par_config': 300,
}
res = executer_dimensionnement(config)
with open('temp_dimensionnement_result.txt', 'w', encoding='utf-8') as out:
    out.write(str(res) + '\n')
