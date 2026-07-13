// G7V1 - Plantilla de transformaciones Power Query (M)
// Ajuste la ruta base según la carpeta donde ubique los archivos CSV.

let
    RutaBase = "C:\\G7V1\\03_data\\",

    ClientesOrigen = Csv.Document(File.Contents(RutaBase & "G7_clientes.csv"),[Delimiter=",", Columns=8, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    ClientesEncabezados = Table.PromoteHeaders(ClientesOrigen, [PromoteAllScalars=true]),
    Dim_Clientes = Table.TransformColumnTypes(ClientesEncabezados,{
        {"ID_Cliente", type text}, {"Ciudad", type text}, {"Segmento", type text}, {"Edad", Int64.Type},
        {"IngresoMensual", Int64.Type}, {"CanalPreferido", type text}, {"FechaAlta", type date}, {"EstadoCliente", type text}
    }),

    VentasOrigen = Csv.Document(File.Contents(RutaBase & "G7_ventas.csv"),[Delimiter=",", Columns=9, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    VentasEncabezados = Table.PromoteHeaders(VentasOrigen, [PromoteAllScalars=true]),
    VentasTipos = Table.TransformColumnTypes(VentasEncabezados,{
        {"ID_Venta", type text}, {"ID_Cliente", type text}, {"ID_Producto", type text}, {"ID_Canal", type text},
        {"FechaVenta", type date}, {"Cantidad", Int64.Type}, {"ValorUnitario", Int64.Type}, {"Descuento", type number}, {"EstadoVenta", type text}
    }),
    VentasFiltradas = Table.SelectRows(VentasTipos, each [EstadoVenta] = "Completada"),
    VentasBruta = Table.AddColumn(VentasFiltradas, "VentaBruta", each [Cantidad] * [ValorUnitario], Int64.Type),
    VentasDescuento = Table.AddColumn(VentasBruta, "ValorDescuento", each Number.Round([VentaBruta] * [Descuento],0), Int64.Type),
    Fact_Ventas = Table.AddColumn(VentasDescuento, "VentaNeta", each [VentaBruta] - [ValorDescuento], Int64.Type)
in
    Fact_Ventas