
# ============================================================
# G6V1 - Integración de datos desde CSV, Excel, JSON y SQL relacional
# ============================================================

from pathlib import Path
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt

# ------------------------------------------------------------
# 1. Configuración de rutas
# ------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
DATA_DIR = BASE_DIR.parent / "03_data" if (BASE_DIR.parent / "03_data").exists() else BASE_DIR / "03_data"
OUT_DIR = BASE_DIR.parent / "08_revision" / "evidencias" if (BASE_DIR.parent / "08_revision").exists() else BASE_DIR / "salidas"
OUT_DIR.mkdir(parents=True, exist_ok=True)

print("Carpeta de datos:", DATA_DIR)
print("Carpeta de salidas:", OUT_DIR)

# ------------------------------------------------------------
# 2. Carga de fuentes
# ------------------------------------------------------------
clientes = pd.read_csv(DATA_DIR / "clientes.csv")

try:
    compras = pd.read_excel(DATA_DIR / "compras.xlsx", sheet_name="compras")
except Exception as e:
    print("No se pudo leer compras.xlsx. Se usará compras_respaldo.csv. Detalle:", e)
    compras = pd.read_csv(DATA_DIR / "compras_respaldo.csv")

web = pd.read_json(DATA_DIR / "interacciones_web.json")

conn = sqlite3.connect(DATA_DIR / "empresa_integracion.db")
satisfaccion = pd.read_sql_query("SELECT * FROM satisfaccion_clientes", conn)
conn.close()

print("clientes:", clientes.shape)
print("compras:", compras.shape)
print("web:", web.shape)
print("satisfaccion:", satisfaccion.shape)

# ------------------------------------------------------------
# 3. Diagnóstico inicial
# ------------------------------------------------------------
fuentes = {
    "clientes": clientes,
    "compras": compras,
    "web": web,
    "satisfaccion": satisfaccion
}

for nombre, df in fuentes.items():
    print("\nFuente:", nombre)
    print("Dimensiones:", df.shape)
    print("Duplicados:", df.duplicated().sum())
    print("Nulos por columna:")
    print(df.isna().sum())

# ------------------------------------------------------------
# 4. Normalización de llaves y formatos
# ------------------------------------------------------------
for df in [clientes, compras, web, satisfaccion]:
    df["ID_Cliente"] = df["ID_Cliente"].astype(str).str.strip().str.upper()

clientes["FechaRegistro"] = pd.to_datetime(clientes["FechaRegistro"], errors="coerce")
compras["FechaCompra"] = pd.to_datetime(compras["FechaCompra"], errors="coerce")

# ------------------------------------------------------------
# 5. Validación de correspondencia entre fuentes
# ------------------------------------------------------------
ids_clientes = set(clientes["ID_Cliente"])
ids_compras = set(compras["ID_Cliente"])

compras_sin_cliente = compras[~compras["ID_Cliente"].isin(ids_clientes)]
print("\nCompras con ID_Cliente que no existe en clientes:", len(compras_sin_cliente))
if len(compras_sin_cliente) > 0:
    print(compras_sin_cliente[["ID_Cliente", "FechaCompra", "CategoriaProducto"]].head())

# ------------------------------------------------------------
# 6. Agregación de compras por cliente
# ------------------------------------------------------------
compras["ValorNeto"] = compras["Cantidad"] * compras["ValorUnitario"] * (1 - compras["Descuento"])

compras_validas = compras[compras["ID_Cliente"].isin(ids_clientes)].copy()

compras_agregadas = (
    compras_validas
    .groupby("ID_Cliente")
    .agg(
        TotalCompras=("ID_Cliente", "count"),
        ValorTotalCompras=("ValorNeto", "sum"),
        TicketPromedio=("ValorNeto", "mean"),
        UltimaCompra=("FechaCompra", "max")
    )
    .reset_index()
)

compras_agregadas["ValorTotalCompras"] = compras_agregadas["ValorTotalCompras"].round(2)
compras_agregadas["TicketPromedio"] = compras_agregadas["TicketPromedio"].round(2)

print("\nCompras agregadas:")
print(compras_agregadas.head())

# ------------------------------------------------------------
# 7. Integración de fuentes
# ------------------------------------------------------------
dataset = clientes.merge(compras_agregadas, on="ID_Cliente", how="left")
dataset = dataset.merge(web, on="ID_Cliente", how="left")
dataset = dataset.merge(satisfaccion, on="ID_Cliente", how="left")

# Rellenar variables de compras para clientes sin compras
dataset["TotalCompras"] = dataset["TotalCompras"].fillna(0).astype(int)
dataset["ValorTotalCompras"] = dataset["ValorTotalCompras"].fillna(0)
dataset["TicketPromedio"] = dataset["TicketPromedio"].fillna(0)

# ------------------------------------------------------------
# 8. Validación del dataset integrado
# ------------------------------------------------------------
print("\nDataset integrado:", dataset.shape)
print("Duplicados por ID_Cliente:", dataset["ID_Cliente"].duplicated().sum())
print("Nulos finales:")
print(dataset.isna().sum())

# ------------------------------------------------------------
# 9. Indicadores básicos
# ------------------------------------------------------------
indicadores_segmento = (
    dataset.groupby("Segmento")
    .agg(
        Clientes=("ID_Cliente", "count"),
        ValorTotal=("ValorTotalCompras", "sum"),
        TicketPromedio=("TicketPromedio", "mean"),
        QuejasPromedio=("QuejasUltimos6M", "mean")
    )
    .round(2)
    .reset_index()
)

indicadores_ciudad = (
    dataset.groupby("Ciudad")
    .agg(
        Clientes=("ID_Cliente", "count"),
        ValorTotal=("ValorTotalCompras", "sum"),
        VisitasPromedio=("VisitasWebUltimoMes", "mean")
    )
    .round(2)
    .reset_index()
)

print("\nIndicadores por segmento:")
print(indicadores_segmento)

print("\nIndicadores por ciudad:")
print(indicadores_ciudad)

# ------------------------------------------------------------
# 10. Visualizaciones
# ------------------------------------------------------------
plt.figure(figsize=(8, 5))
indicadores_segmento.plot(kind="bar", x="Segmento", y="ValorTotal", legend=False)
plt.title("Valor total de compras por segmento")
plt.xlabel("Segmento")
plt.ylabel("Valor total")
plt.tight_layout()
plt.savefig(OUT_DIR / "valor_total_por_segmento.png", dpi=150)
plt.show()

plt.figure(figsize=(8, 5))
indicadores_ciudad.plot(kind="bar", x="Ciudad", y="ValorTotal", legend=False)
plt.title("Valor total de compras por ciudad")
plt.xlabel("Ciudad")
plt.ylabel("Valor total")
plt.tight_layout()
plt.savefig(OUT_DIR / "valor_total_por_ciudad.png", dpi=150)
plt.show()

plt.figure(figsize=(8, 5))
dataset["Satisfaccion"].value_counts().plot(kind="bar")
plt.title("Distribución de satisfacción")
plt.xlabel("Satisfacción")
plt.ylabel("Cantidad de clientes")
plt.tight_layout()
plt.savefig(OUT_DIR / "distribucion_satisfaccion.png", dpi=150)
plt.show()

# ------------------------------------------------------------
# 11. Exportación de resultados
# ------------------------------------------------------------
dataset.to_csv(OUT_DIR / "dataset_integrado_clientes.csv", index=False, encoding="utf-8-sig")
indicadores_segmento.to_csv(OUT_DIR / "indicadores_segmento.csv", index=False, encoding="utf-8-sig")
indicadores_ciudad.to_csv(OUT_DIR / "indicadores_ciudad.csv", index=False, encoding="utf-8-sig")

reporte = []
reporte.append("REPORTE DE VALIDACIÓN G6V1")
reporte.append("=" * 40)
reporte.append(f"Clientes fuente: {len(clientes)}")
reporte.append(f"Compras fuente: {len(compras)}")
reporte.append(f"Compras sin cliente: {len(compras_sin_cliente)}")
reporte.append(f"Filas dataset integrado: {len(dataset)}")
reporte.append(f"Columnas dataset integrado: {dataset.shape[1]}")
reporte.append(f"Duplicados por ID_Cliente en dataset final: {dataset['ID_Cliente'].duplicated().sum()}")
reporte.append("\nNulos finales:")
reporte.append(str(dataset.isna().sum()))

with open(OUT_DIR / "reporte_validacion_integracion.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(reporte))

print("\nArchivos generados en:", OUT_DIR)
print("Proceso finalizado correctamente.")
