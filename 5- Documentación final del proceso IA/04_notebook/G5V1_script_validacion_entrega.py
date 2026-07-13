import pandas as pd
from pathlib import Path

ruta = Path('G5V1_dataset_final_segmentado.csv')
df = pd.read_csv(ruta)
resumen = pd.DataFrame({
    'Indicador':['Filas','Columnas','Duplicados','Nulos totales','Tiene Cluster','Tiene Abandono'],
    'Resultado':[len(df), df.shape[1], df.duplicated().sum(), df.isnull().sum().sum(), 'Sí' if 'Cluster' in df.columns else 'No', 'Sí' if 'Abandono' in df.columns else 'No']
})
print(resumen)
resumen.to_csv('G5V1_resumen_validacion_generado.csv', index=False)
if 'Cluster' in df.columns:
    perfil = df.groupby('Cluster').mean(numeric_only=True).round(2)
    print(perfil)
    perfil.to_csv('G5V1_perfil_clusters_generado.csv')