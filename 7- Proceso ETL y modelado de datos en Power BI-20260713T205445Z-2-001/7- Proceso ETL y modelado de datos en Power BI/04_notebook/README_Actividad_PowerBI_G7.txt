G7V1 - Actividad práctica en Power BI

Objetivo: realizar un proceso ETL con Power Query y construir un modelo de datos básico para analítica.

Orden recomendado:
1. Abrir Power BI Desktop.
2. Obtener datos desde carpeta o archivos CSV ubicados en 03_data.
3. Cargar clientes, ventas, productos, canales, tickets_soporte y metas_segmento.
4. En Power Query: promover encabezados, cambiar tipos, limpiar datos, filtrar ventas completadas y crear columnas VentaBruta, ValorDescuento y VentaNeta.
5. Crear consultas de dimensiones y hechos: Dim_Clientes, Dim_Productos, Dim_Canales, Fact_Ventas, Fact_Soporte y Metas_Segmento.
6. Cerrar y aplicar.
7. Crear relaciones: Fact_Ventas con Dim_Clientes, Dim_Productos, Dim_Canales y Dim_Fecha.
8. Crear tabla calendario con DAX.
9. Crear medidas DAX básicas usando G7_medidas_DAX.txt.
10. Validar totales contra los archivos de referencia del docente o contra los totales del informe.
11. Guardar el archivo .pbix y exportar evidencias de modelo, Power Query y medidas.

Nota: el archivo PBIX lo construye el aprendiz en Power BI Desktop. Este paquete contiene los datos, guías y scripts necesarios.