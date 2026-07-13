
import pandas as pd
from pathlib import Path

base = Path(__file__).resolve().parent.parent / '03_data'
clientes = pd.read_csv(base / 'G7_clientes.csv')
ventas = pd.read_csv(base / 'G7_ventas.csv')
productos = pd.read_csv(base / 'G7_productos.csv')
canales = pd.read_csv(base / 'G7_canales.csv')

fact = ventas[ventas['EstadoVenta'] == 'Completada'].copy()
fact['VentaBruta'] = fact['Cantidad'] * fact['ValorUnitario']
fact['ValorDescuento'] = (fact['VentaBruta'] * fact['Descuento']).round(0).astype(int)
fact['VentaNeta'] = fact['VentaBruta'] - fact['ValorDescuento']

print('Registros clientes:', len(clientes))
print('Registros ventas originales:', len(ventas))
print('Registros ventas completadas:', len(fact))
print('Total venta neta:', int(fact['VentaNeta'].sum()))
print('Duplicados ID_Cliente:', clientes['ID_Cliente'].duplicated().sum())
print('Duplicados ID_Producto:', productos['ID_Producto'].duplicated().sum())
print('Duplicados ID_Canal:', canales['ID_Canal'].duplicated().sum())

fact.to_csv(base / 'G7_fact_ventas_validada_desde_python.csv', index=False, encoding='utf-8-sig')
