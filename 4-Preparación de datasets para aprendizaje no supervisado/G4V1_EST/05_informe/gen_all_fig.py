import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import StandardScaler, OrdinalEncoder, OneHotEncoder
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.decomposition import PCA

OUT = r'C:\Users\SENA\Documents\VS_CODE\Sena_IA_1\IA_1\4-Preparación de datasets para aprendizaje no supervisado\G4V1_EST\05_informe'

# ── Load data ──
df = pd.read_csv(r'C:\Users\SENA\Documents\VS_CODE\Sena_IA_1\IA_1\4-Preparación de datasets para aprendizaje no supervisado\G4V1_EST\04_notebook\G4_base_clientes.csv')

variables_cluster = [
    'Edad', 'IngresoMensual', 'CantidadCompras', 'ComprasUltimos12M',
    'AntiguedadMeses', 'QuejasUltimos6M', 'DiasDesdeUltimaCompra',
    'VisitasWebUltimoMes', 'TiempoPromedioSesionMin', 'CuponesUsados',
    'Ciudad', 'CanalPreferido', 'ZonaResidencia', 'Segmento', 'Satisfaccion'
]
X_cluster = df[variables_cluster].copy()

columnas_numericas = [
    'Edad', 'IngresoMensual', 'CantidadCompras', 'ComprasUltimos12M',
    'AntiguedadMeses', 'QuejasUltimos6M', 'DiasDesdeUltimaCompra',
    'VisitasWebUltimoMes', 'TiempoPromedioSesionMin', 'CuponesUsados'
]
columnas_nominales = ['Ciudad', 'CanalPreferido', 'ZonaResidencia']
columnas_ordinales = ['Segmento', 'Satisfaccion']

try:
    onehot = OneHotEncoder(handle_unknown='ignore', sparse_output=False)
except TypeError:
    onehot = OneHotEncoder(handle_unknown='ignore', sparse=False)

preprocesador_cluster = ColumnTransformer(
    transformers=[
        ('num', StandardScaler(), columnas_numericas),
        ('nom', onehot, columnas_nominales),
        ('ord', OrdinalEncoder(categories=[['Basico', 'Medio', 'Premium'], ['Baja', 'Media', 'Alta']], handle_unknown='use_encoded_value', unknown_value=-1), columnas_ordinales)
    ]
)
X_preparado = preprocesador_cluster.fit_transform(X_cluster)

AZUL = '#1B3A6B'
GRIS = '#808080'
NEGRO = '#000000'
C0, C1, C2 = '#1B3A6B', '#808080', '#000000'

# ═══════ FIG 1: CODO (2-10) ═══════
rango_k = range(2, 11)
inercias = []
for k in rango_k:
    m = KMeans(n_clusters=k, random_state=42, n_init=10)
    m.fit(X_preparado); inercias.append(m.inertia_)
tc = pd.DataFrame({'k': list(rango_k), 'inercia': inercias})
tc['ri'] = tc['inercia'].shift(1) - tc['inercia']
tc['rp'] = (tc['ri'] / tc['inercia'].shift(1) * 100).round(2)
plt.figure(figsize=(8.5,5))
plt.plot(list(rango_k), inercias, marker='o', color=AZUL, lw=2.5, ms=8)
plt.title('Metodo del Codo (k=2 a 10)', fontsize=14, fontweight='bold', color=AZUL)
plt.xlabel('k'); plt.ylabel('Inercia'); plt.grid(True, ls='--', alpha=.5, color=GRIS)
plt.xticks(list(rango_k))
for i, v in enumerate(inercias):
    plt.text(list(rango_k)[i]+0.15, v+30, f'{v:.0f}', fontsize=8, color=AZUL)
plt.tight_layout(); plt.savefig(f'{OUT}\\fig_codo.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ FIG 2: SILHOUETTE ═══════
sils = []
for k in rango_k:
    m = KMeans(n_clusters=k, random_state=42, n_init=10)
    et = m.fit_predict(X_preparado); sils.append(silhouette_score(X_preparado, et))
ts = pd.DataFrame({'k': list(rango_k), 'silhouette_score': sils})
km_sil = int(ts.loc[ts['silhouette_score'].idxmax(), 'k'])
plt.figure(figsize=(8.5,5))
plt.plot(list(rango_k), sils, marker='o', color=AZUL, lw=2.5, ms=8)
plt.title('Silhouette Score (k=2 a 10)', fontsize=14, fontweight='bold', color=AZUL)
plt.xlabel('k'); plt.ylabel('Silhouette Score'); plt.grid(True, ls='--', alpha=.5, color=GRIS)
plt.xticks(list(rango_k))
for i, v in enumerate(sils):
    plt.text(list(rango_k)[i]+0.15, v+0.002, f'{v:.4f}', fontsize=8, color=AZUL)
plt.tight_layout(); plt.savefig(f'{OUT}\\fig_silhouette.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ MODELO FINAL ═══════
k_opt = 3
modelo = KMeans(n_clusters=k_opt, random_state=42, n_init=10)
clusters = modelo.fit_predict(X_preparado)
df['Cluster'] = clusters

# ═══════ FIG 3: DISPERSION SIN ═══════
plt.figure(figsize=(8,5.5))
plt.scatter(df['IngresoMensual'], df['CantidadCompras'], alpha=.7, color=AZUL, s=30)
plt.title('Dispersion: Ingreso mensual vs Cantidad de compras', fontsize=13, fontweight='bold', color=AZUL)
plt.xlabel('Ingreso mensual'); plt.ylabel('Cantidad de compras'); plt.grid(True, ls='--', alpha=.5, color=GRIS)
plt.tight_layout(); plt.savefig(f'{OUT}\\fig_disp_sin.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ FIG 4: DISPERSION CON ═══════
plt.figure(figsize=(8,5.5))
sc = plt.scatter(df['IngresoMensual'], df['CantidadCompras'], c=df['Cluster'], alpha=.75, cmap='viridis', s=30)
plt.title('Dispersion con clusters: Ingreso vs Compras', fontsize=13, fontweight='bold', color=AZUL)
plt.xlabel('Ingreso mensual'); plt.ylabel('Cantidad de compras'); plt.grid(True, ls='--', alpha=.5, color=GRIS)
plt.legend(*sc.legend_elements(), title='Cluster'); plt.tight_layout()
plt.savefig(f'{OUT}\\fig_disp_con.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ FIG 5: PCA ═══════
pca = PCA(n_components=2, random_state=42)
comps = pca.fit_transform(X_preparado)
ve = pca.explained_variance_ratio_
plt.figure(figsize=(8.5,6))
for i, (c, lab) in enumerate([(C0,'Cluster 0'),(GRIS,'Cluster 1'),(NEGRO,'Cluster 2')]):
    mask = clusters == i
    plt.scatter(comps[mask,0], comps[mask,1], c=c, label=lab, alpha=.7, s=35)
plt.title(f'PCA: {ve[0]*100:.1f}% + {ve[1]*100:.1f}% = {ve.sum()*100:.1f}% varianza', fontsize=13, fontweight='bold', color=AZUL)
plt.xlabel(f'PC1 ({ve[0]*100:.1f}%)'); plt.ylabel(f'PC2 ({ve[1]*100:.1f}%)')
plt.grid(True, ls='--', alpha=.5, color=GRIS); plt.legend()
plt.tight_layout(); plt.savefig(f'{OUT}\\fig_pca.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ FIG 6: BARRAS ═══════
pn = df.groupby('Cluster')[columnas_numericas].mean()
vg = ['AntiguedadMeses','CantidadCompras','QuejasUltimos6M','DiasDesdeUltimaCompra','CuponesUsados']
ng = ['Antiguedad','Compras\nTotales','Quejas','Dias sin\ncompra','Cupones']
x = np.arange(len(vg)); w = .25
fig, ax = plt.subplots(figsize=(10,5.5))
for i, col in enumerate([C0,GRIS,NEGRO]):
    ax.bar(x+i*w, [pn.loc[i,v] for v in vg], w, label=f'Cluster {i}', color=col, alpha=.85)
ax.set_title('Perfil Comparativo de Clusters', fontsize=14, fontweight='bold', color=AZUL)
ax.set_xticks(x+w); ax.set_xticklabels(ng, fontsize=9)
ax.set_ylabel('Valor Promedio'); ax.legend(); ax.grid(axis='y', ls='--', alpha=.5, color=GRIS)
plt.tight_layout(); plt.savefig(f'{OUT}\\fig_barras.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ FIG 7: BOXPLOTS ═══════
fig, axes = plt.subplots(2, 3, figsize=(13, 8))
vars_box = ['CantidadCompras', 'AntiguedadMeses', 'QuejasUltimos6M', 'DiasDesdeUltimaCompra', 'IngresoMensual', 'VisitasWebUltimoMes']
titles_box = ['Compras Totales', 'Antiguedad (meses)', 'Quejas (6m)', 'Dias sin compra', 'Ingreso Mensual', 'Visitas Web']
for ax, var, tit in zip(axes.flat, vars_box, titles_box):
    data = [df[df['Cluster']==i][var] for i in range(3)]
    bp = ax.boxplot(data, patch_artist=True, widths=0.5)
    for patch, color in zip(bp['boxes'], ['#1B3A6B', '#808080', '#000000']):
        patch.set_facecolor(color); patch.set_alpha(0.6)
    ax.set_title(tit, fontsize=10, fontweight='bold', color=AZUL)
    ax.set_xticklabels(['C0','C1','C2']); ax.grid(True, ls='--', alpha=.3)
plt.suptitle('Distribucion de Variables por Cluster', fontsize=14, fontweight='bold', color=AZUL)
plt.tight_layout(); plt.savefig(f'{OUT}\\fig_boxplots.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ FIG 8: CORRELACION HEATMAP ═══════
corr = df[columnas_numericas + ['Cluster']].corr()
fig, ax = plt.subplots(figsize=(10, 8))
im = ax.imshow(corr, cmap='Blues', aspect='auto')
ax.set_xticks(range(len(corr.columns))); ax.set_yticks(range(len(corr.columns)))
ax.set_xticklabels(corr.columns, rotation=45, ha='right', fontsize=8)
ax.set_yticklabels(corr.columns, fontsize=8)
for i in range(len(corr.columns)):
    for j in range(len(corr.columns)):
        ax.text(j, i, f'{corr.iloc[i,j]:.2f}', ha='center', va='center', fontsize=7, color='white' if abs(corr.iloc[i,j])>0.4 else 'black')
plt.title('Matriz de Correlacion', fontsize=14, fontweight='bold', color=AZUL)
plt.tight_layout(); plt.savefig(f'{OUT}\\fig_correlacion.png', dpi=200, bbox_inches='tight'); plt.close()

# ═══════ EXPORT DATA ═══════
pn.round(2).to_csv(f'{OUT}\\perfil_numerico.csv')
def moda_segura(s): m = s.mode(); return np.nan if len(m)==0 else m.iloc[0]
pc = df.groupby('Cluster')[columnas_nominales+columnas_ordinales].agg(moda_segura)
pc.to_csv(f'{OUT}\\perfil_categorico.csv')
ra = df.groupby('Cluster')['Abandono'].agg(['count','mean'])
ra['mean'] = (ra['mean']*100).round(2)
ra.to_csv(f'{OUT}\\resumen_abandono.csv')
tc.round(2).merge(ts.round(4), on='k').to_csv(f'{OUT}\\tabla_comparativa_k.csv', index=False)
# Estadisticas descriptivas por cluster
stats = df.groupby('Cluster')[columnas_numericas].describe().round(2)
stats.to_csv(f'{OUT}\\estadisticas_clusters.csv')
# Conteo de categorias
for col in columnas_nominales+columnas_ordinales:
    ct = pd.crosstab(df['Cluster'], df[col])
    ct.to_csv(f'{OUT}\\frecuencia_{col}.csv')

print(f"Figuras generadas: codo, silhouette, disp_sin, disp_con, pca, barras, boxplots, correlacion")
print(f"Datos exportados a {OUT}")
print(f"Mejor k por silhouette: {km_sil}")
print(f"Varianza PCA: PC1={ve[0]*100:.2f}%, PC2={ve[1]*100:.2f}%, Total={ve.sum()*100:.2f}%")
