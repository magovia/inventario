Attribute VB_Name = "appFacturaDetalle"
Option Compare Database
Option Explicit

Public Function AppendFacturasDetalleUsingRecordset()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsDestination As DAO.Recordset
    Dim intRecordsAffected As Integer
    
    ' Get reference to current database
    Set db = CurrentDb
    
    ' Open source recordset
    Set rsSource = db.OpenRecordset("tblLineaDetalle", dbOpenSnapshot)
    
    ' Open destination recordset
    Set rsDestination = db.OpenRecordset("tblFacturasDetalle", dbOpenDynaset)
    
    ' Check if source recordset has records
    If rsSource.EOF Then
        MsgBox "No se encontraron registros en la tabla tblLineaDetalle.", vbInformation, "No hay registros para anexar"
        GoTo CleanUp
    End If
    
    ' Loop through source recordset
    rsSource.MoveFirst
    intRecordsAffected = 0
    
    Do Until rsSource.EOF
        ' Add new record to destination
        rsDestination.AddNew
        
        ' Map the fields from source to destination
        rsDestination!FacturaID = rsSource!NumeroConsecutivo
        rsDestination!Codigo = rsSource!Codigo
        rsDestination!IdProducto = rsSource!IdProducto
        rsDestination!Cantidad = rsSource!Cantidad
        rsDestination!PrecioUnitario = rsSource!PrecioUnitario
        rsDestination!MontoTot = rsSource!MontoTotal
        rsDestination!ImpuestoPorcentaje = rsSource!PorcentajeImpuesto
        rsDestination!CodigoImpuesto = rsSource!CodImpuesto
        rsDestination!ImpuestoCodTarifa = rsSource!CodigoTarifa
        rsDestination!NaturalezaDescuento = rsSource!NaturalezaDescuento
        rsDestination!Descuento = rsSource!Descuento
        rsDestination!Impuesto = rsSource!Impuesto
        rsDestination!SubTotal = rsSource!SubTotal
        rsDestination!MontoTotalLinea = rsSource!MontoTotalLinea
        
        ' Save the new record
        rsDestination.Update
        
        ' Increment counter
        intRecordsAffected = intRecordsAffected + 1
        
        ' Move to next record
        rsSource.MoveNext
    Loop
    
    ' Report success
    MsgBox intRecordsAffected & " registros agregados a la tabla tblFacturasDetalle.", _
           vbInformation, "Proceso de agregar"
    
CleanUp:
    ' Clean up
    If Not rsSource Is Nothing Then
        rsSource.Close
    End If
    If Not rsDestination Is Nothing Then
        rsDestination.Close
    End If
    Set rsSource = Nothing
    Set rsDestination = Nothing
    Set db = Nothing
    AppendFacturasDetalleUsingRecordset = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in AppendFacturasDetalleUsingRecordset"
    
    ' Close recordsets if they are open
    If Not rsSource Is Nothing Then
        rsSource.Close
    End If
    If Not rsDestination Is Nothing Then
        rsDestination.Close
    End If
    
    Set rsSource = Nothing
    Set rsDestination = Nothing
    Set db = Nothing
    AppendFacturasDetalleUsingRecordset = False
End Function

