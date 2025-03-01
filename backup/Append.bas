Attribute VB_Name = "Append"
Option Compare Database
Option Explicit

Public Sub RellenarLineaDetalleFromXLS()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim sqlInsert As String

    ' Get the current database
    Set db = CurrentDb()
    
    ' SQL to aggregate data from XLS_Detalle and append into tblLineaDetalle
    sqlInsert = _
        "INSERT INTO tblLineaDetalle " & _
        "(IdProducto, NumeroConsecutivo, Codigo, UnidadMedida, Cantidad, Detalle, " & _
        "PrecioUnitario, PorcentajeImpuesto, Descuento, SubTotal, Impuesto, MontoTotal, MontoTotalLinea) " & _
        "SELECT " & _
        "XLS_Detalle.id_producto AS IdProducto, " & _
        "XLS_Detalle.FacturaID AS NumeroConsecutivo, " & _
        "XLS_Detalle.Codigo, " & _
        "XLS_Detalle.UnidadMedida, " & _
        "Sum(XLS_Detalle.Cantidad) AS Cantidad, " & _
        "XLS_Detalle.Detalle, " & _
        "Sum(XLS_Detalle.PrecioUnitario) AS PrecioUnitario, " & _
        "XLS_Detalle.ImpuestoPorcentaje AS PorcentajeImpuesto, " & _
        "Sum(XLS_Detalle.MontoDescuento) AS Descuento, " & _
        "Sum(XLS_Detalle.SubTotal) AS SubTotal, " & _
        "Sum(XLS_Detalle.MontoImpuesto) AS Impuesto, " & _
        "Sum(XLS_Detalle.MontoTotal) AS MontoTotal, " & _
        "Sum(XLS_Detalle.MontoTotalLinea) AS MontoTotalLinea " & _
        "FROM XLS_Detalle " & _
        "GROUP BY XLS_Detalle.id_producto, XLS_Detalle.FacturaID, XLS_Detalle.Codigo, " & _
        "XLS_Detalle.UnidadMedida, XLS_Detalle.Detalle, XLS_Detalle.ImpuestoPorcentaje;"

    ' Execute the SQL statement
    db.Execute sqlInsert, dbFailOnError
    MsgBox "Data appended successfully from XLS_Detalle to tblLineaDetalle.", vbInformation, "Success"
    
ExitHandler:
    ' Cleanup
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    Resume ExitHandler
End Sub


Sub execcuteQueries()

Dim db As DAO.Database
Set db = CurrentDb

    On Error GoTo ErrorHandler
    
    
   ' Db.Execute "SELECT qryImpuesto.* INTO tblImpuestos FROM qryImpuesto;", dbFailOnError
    
    'Debug.Print Db.RecordsAffected
    
    'Append dtos en tblProveedores
    db.Execute "INSERT INTO tblProveedor SELECT qryProveedor.* FROM qryProveedor;", dbFailOnError
    
    Debug.Print db.RecordsAffected
    
    'Append datos en tblProductos
    db.Execute "INSERT INTO tblProductos ( idProducto, Detalle, PrecioUnitario, UnidadMedida, UnidadMedidaComercial, Comentario )" & _
    " SELECT qryProducto.IdProducto, qryProducto.Detalle, qryProducto.PrecioUnitario, qryProducto.UnidadMedida, qryProducto.UnidadMedidaComercial, qryProducto.Comentario" & _
    " FROM qryProducto;", dbFailOnError
    
    Debug.Print db.RecordsAffected
        
    db.Execute "INSERT INTO tblLineaDetalle ( NumeroLinea, IdProducto, NumeroConsecutivo, Codigo, Cantidad, UnidadMedida, UnidadMedidaComercial, CodImpuesto, CodigoTarifa, Detalle, PorcentajeImpuesto, PrecioUnitario, Descuento, SubTotal, Impuesto, MontoTot, MontoTotalLinea )" & _
    " SELECT LineaDetalle.NumeroLinea, Left([LineaDetalle]![Codigo],5) & " - " & Len([Detalle]) AS IdProducto, FacturaElectronica.NumeroConsecutivo AS FacturaID, LineaDetalle.Codigo, FormatNumber([Cantidad],2) AS Cant, LineaDetalle.UnidadMedida, LineaDetalle.UnidadMedidaComercial, tblImpuestos.CodImpuesto, tblImpuestos.CodigoTarifa, LineaDetalle.Detalle, tblImpuestos.PorcentajeImpuesto, FormatNumber([PrecioUnitario],2) AS PrecioUnit, [LineaDetalle]![MontoTotal]-[LineaDetalle]![SubTotal] AS Descuento, FormatNumber([SubTotal],2) AS SubTot, tblImpuestos.MontoImpuesto, FormatNumber([MontoTotal],2) AS MontoTot, FormatNumber([MontoTotalLinea],2) AS MontoTotalLin " & _
    " FROM FacturaElectronica, LineaDetalle INNER JOIN tblImpuestos ON LineaDetalle.NumeroLinea = tblImpuestos.auto;", dbFailOnError
   
   Debug.Print db.RecordsAffected
     
    'Append datos en tblFacturas
    db.Execute "INSERT INTO tblFacturas ( facturaID, FechaEmision, IdPersoneria, CondicionVenta, PlazoCredito, MedioPago ) " & _
                " SELECT qryFacturas.facturaID, qryFacturas.Fecha, qryFacturas.IdPersoneria, qryFacturas.CondicionVenta, qryFacturas.PlazoCredito, qryFacturas.MedioPago " & _
                " FROM qryFacturas;", dbFailOnError
    
    Debug.Print db.RecordsAffected
                
    'append datos en tblFacturasDetalle
    db.Execute "INSERT INTO tblFacturasDetalle ( IdProducto, FacturaID, Codigo, Cantidad, ImpuestoPorcentaje, NaturalezaDescuento, CodigoImpuesto, ImpuestoCodTarifa, PrecioUnitario, Descuento, SubTotal, Impuesto, MontoTotalLinea ) " & _
                " SELECT tblLineaDetalle.IdProducto, tblLineaDetalle.NumeroConsecutivo, tblLineaDetalle.Codigo, tblLineaDetalle.Cantidad, tblLineaDetalle.PorcentajeImpuesto, tblLineaDetalle.NaturalezaDescuento, tblLineaDetalle.CodImpuesto, tblLineaDetalle.CodigoTarifa, tblLineaDetalle.PrecioUnitario, tblLineaDetalle.Descuento, tblLineaDetalle.SubTotal, tblLineaDetalle.Impuesto, tblLineaDetalle.MontoTotalLinea " & _
                " FROM tblLineaDetalle;", dbFailOnError

    Debug.Print db.RecordsAffected
    
    Debug.Print "All records have been successfully appended from the tables."

    Exit Sub
    
ErrorHandler:
    ' Re-enable warnings in case of an error
    DoCmd.SetWarnings True
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub
Sub AppendTables()
    On Error GoTo ErrorHandler
    
    ' Disable warnings to avoid confirmation prompts
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "SELECT qryImpuesto.* INTO tblImpuestos FROM qryImpuesto;"
    
    'Append dtos en tblProveedores
    DoCmd.RunSQL "INSERT INTO tblProveedor SELECT qryProveedor.* FROM qryProveedor;"
    
    'Append datos en tblProductos
    DoCmd.RunSQL "INSERT INTO tblProductos ( idProducto, Detalle, PrecioUnitario, UnidadMedida, UnidadMedidaComercial, Comentario )" & _
    " SELECT qryProducto.IdProducto, qryProducto.Detalle, qryProducto.PrecioUnitario, qryProducto.UnidadMedida, qryProducto.UnidadMedidaComercial, qryProducto.Comentario" & _
    " FROM qryProducto;"
    
        
    DoCmd.RunSQL "INSERT INTO tblLineaDetalle ( NumeroLinea, IdProducto, NumeroConsecutivo, Codigo, Cantidad, UnidadMedida, UnidadMedidaComercial, CodImpuesto, CodigoTarifa, Detalle, PorcentajeImpuesto, PrecioUnitario, Descuento, SubTotal, Impuesto, MontoTot, MontoTotalLinea )" & _
    " SELECT LineaDetalle.NumeroLinea, Left([LineaDetalle]![Codigo],5) & " - " & Len([Detalle]) AS IdProducto, FacturaElectronica.NumeroConsecutivo AS FacturaID, LineaDetalle.Codigo, FormatNumber([Cantidad],2) AS Cant, LineaDetalle.UnidadMedida, LineaDetalle.UnidadMedidaComercial, tblImpuestos.CodImpuesto, tblImpuestos.CodigoTarifa, LineaDetalle.Detalle, tblImpuestos.PorcentajeImpuesto, FormatNumber([PrecioUnitario],2) AS PrecioUnit, [LineaDetalle]![MontoTotal]-[LineaDetalle]![SubTotal] AS Descuento, FormatNumber([SubTotal],2) AS SubTot, tblImpuestos.MontoImpuesto, FormatNumber([MontoTotal],2) AS MontoTot, FormatNumber([MontoTotalLinea],2) AS MontoTotalLin " & _
    " FROM FacturaElectronica, LineaDetalle INNER JOIN tblImpuestos ON LineaDetalle.NumeroLinea = tblImpuestos.auto;"
   
     
    'Append datos en tblFacturas
    DoCmd.RunSQL "INSERT INTO tblFacturas ( facturaID, FechaEmision, IdPersoneria, CondicionVenta, PlazoCredito, MedioPago ) " & _
                " SELECT qryFacturas.facturaID, qryFacturas.Fecha, qryFacturas.IdPersoneria, qryFacturas.CondicionVenta, qryFacturas.PlazoCredito, qryFacturas.MedioPago " & _
                " FROM qryFacturas;"

                
    'append datos en tblFacturasDetalle
    DoCmd.RunSQL "INSERT INTO tblFacturasDetalle ( IdProducto, FacturaID, Codigo, Cantidad, ImpuestoPorcentaje, NaturalezaDescuento, CodigoImpuesto, ImpuestoCodTarifa, PrecioUnitario, Descuento, SubTotal, Impuesto, MontoTotalLinea ) " & _
                " SELECT tblLineaDetalle.IdProducto, tblLineaDetalle.NumeroConsecutivo, tblLineaDetalle.Codigo, tblLineaDetalle.Cantidad, tblLineaDetalle.PorcentajeImpuesto, tblLineaDetalle.NaturalezaDescuento, tblLineaDetalle.CodImpuesto, tblLineaDetalle.CodigoTarifa, tblLineaDetalle.PrecioUnitario, tblLineaDetalle.Descuento, tblLineaDetalle.SubTotal, tblLineaDetalle.Impuesto, tblLineaDetalle.MontoTotalLinea " & _
                " FROM tblLineaDetalle;"


    ' Re-enable warnings
    DoCmd.SetWarnings True
    
    Debug.Print "All records have been successfully appended from the tables."

    Exit Sub
    
ErrorHandler:
    ' Re-enable warnings in case of an error
    DoCmd.SetWarnings True
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub


Function RunActionQuery(Qry As String, msg As String)

Dim db As DAO.Database
Set db = CurrentDb

    'To Catch the Query Error use dbFailOnError option
    On Error GoTo ErrorHandler
    'DoCmd.SetWarnings False
    'On Error Resume Next
        db.Execute Qry, dbFailOnError
        Debug.Print msg & " Registros modificados:(" & db.RecordsAffected & ")"
        
        'MsgBox msg & " Registros modificados:(" & db.RecordsAffected & ")", vbInformation, "Test"
    'DoCmd.SetWarnings True
    Exit Function
ErrorHandler:
    ' Re-enable warnings in case of an error
    DoCmd.SetWarnings True
    Debug.Print Qry & ":Error " & Err.Description
    
End Function
