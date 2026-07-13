import json, sys
sys.stdout = open('C:\\Users\\SENA\\Documents\\VS_CODE\\v2_full.txt', 'w', encoding='utf-8')

path = 'C:\\Users\\SENA\\Documents\\VS_CODE\\Sena_IA_1\\IA_1\\4-Preparaci\u00f3n de datasets para aprendizaje no supervisado\\G4V1_EST\\04_notebook\\G4_notebook_no_supervisado_V2_explicado_dispersion_codo_silhouette.ipynb'

with open(path, 'r', encoding='utf-8') as f:
    nb = json.load(f)

for i, cell in enumerate(nb['cells']):
    src = ''.join(cell['source'])
    src_clean = src.replace('\u2705', '[OK]').replace('\U0001f4a1', '[idea]').replace('\U0001f447', '[down]').replace('\u23f3', '[time]')
    print(f'=== CELL {i} [{cell["cell_type"]}] ===')
    print(src_clean[:2000])
    print()
sys.stdout.close()
