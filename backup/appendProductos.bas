Attribute VB_Name = "appendProductos"
Option Compare Database
Option Explicit

Public Function AppendProductosUsingRecordset()
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
    Set rsDestination = db.OpenRecordset("tblProductos", dbOpenDynaset)
    
    ' Check if source recordset has records
    If rsSource.EOF Then
        MsgBox "No se encontraron registros en tblLineaDetalle.", vbInformation, "Proceso anexar"
        GoTo CleanUp
    End If
    
    ' Loop through source recordset
    rsSource.MoveFirst
    intRecordsAffected = 0
    
    Do Until rsSource.EOF
        ' Add new record to destination
        rsDestination.AddNew
        
        ' Map the fields from source to destination
        rsDestination!IdProducto = rsSource!IdProducto
        rsDestination!CodigoCabys = rsSource!Codigo
        rsDestination!PorcentajeImpuesto = rsSource!PorcentajeImpuesto
        rsDestination!Detalle = rsSource!Detalle
        rsDestination!PrecioUnitario = rsSource!PrecioUnitario
        rsDestination!UnidadMedida = rsSource!UnidadMedida
        rsDestination!UnidadMedidaComercial = rsSource!UnidadMedidaComercial
        rsDestination!Cantidad = 0 ' Hardcoded value as in your SQL
        
        ' Save the new record
        rsDestination.Update
        
        ' Increment counter
        intRecordsAffected = intRecordsAffected + 1
        
        ' Move to next record
        rsSource.MoveNext
    Loop
    
    ' Report success
    MsgBox intRecordsAffected & " Productos han sido registrados.", _
           vbInformation, "Anexar productos"
    
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
    AppendProductosUsingRecordset = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in AppendProductosUsingRecordset"
    
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
    AppendProductosUsingRecordset = False
End Function

' Optional: Create a version with transaction for additional safety
Public Function AppendProductosUsingRecordsetWithTransaction()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rsSource As DAO.Recordset
    Dim rsDestination As DAO.Recordset
    Dim intRecordsAffected As Integer
    
    ' Get reference to workspace and database
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    
    ' Begin transaction
    ws.BeginTrans
    
    ' Open source recordset
    Set rsSource = db.OpenRecordset("tblLineaDetalle", dbOpenSnapshot)
    
    ' Open destination recordset
    Set rsDestination = db.OpenRecordset("tblProductos", dbOpenDynaset)
    
    ' Check if source recordset has records
    If rsSource.EOF Then
        MsgBox "No records found in tblLineaDetalle.", vbInformation, "No Records to Append"
        ws.CommitTrans ' Commit empty transaction
        GoTo CleanUp
    End If
    
    ' Loop through source recordset
    rsSource.MoveFirst
    intRecordsAffected = 0
   
    ' Print header for the debug output
    Debug.Print "--- tblLineaDetalle Records ---"
    Debug.Print "IdProducto | Codigo | PorcentajeImpuesto | Detalle | PrecioUnitario | UnidadMedida | UnidadMedidaComercial"
    Debug.Print String(100, "-") ' Separator line
    
    Do Until rsSource.EOF
    
            ' Debug output to Immediate Window (Ctrl+G to view)
        Debug.Print rsSource!IdProducto & " | " & _
                    rsSource!Codigo & " | " & _
                    rsSource!PorcentajeImpuesto & " | " & _
                    Left(rsSource!Detalle, 20) & "... | " & _
                    rsSource!PrecioUnitario & " | " & _
                    rsSource!UnidadMedida & " | " & _
                    rsSource!UnidadMedidaComercial
                    
        ' Add new record to destination
        rsDestination.AddNew
        
        ' Map the fields from source to destination
        rsDestination!IdProducto = rsSource!IdProducto
        rsDestination!CodigoCabys = rsSource!Codigo
        rsDestination!PorcentajeImpuesto = rsSource!PorcentajeImpuesto
        rsDestination!Detalle = rsSource!Detalle
        rsDestination!PrecioUnitario = rsSource!PrecioUnitario
        rsDestination!UnidadMedida = rsSource!UnidadMedida
        rsDestination!UnidadMedidaComercial = rsSource!UnidadMedidaComercial
        rsDestination!Cantidad = 0 ' Hardcoded value as in your SQL
        
        ' Save the new record
        rsDestination.Update
        
        ' Increment counter
        intRecordsAffected = intRecordsAffected + 1
        
        ' Move to next record
        rsSource.MoveNext
    Loop
    
    ' Commit transaction if all went well
    ws.CommitTrans
    
    ' Report success
    MsgBox intRecordsAffected & " Productos han sido guardados.", _
           vbInformation, "Proceso anexar productos"
    
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
    Set ws = Nothing
    AppendProductosUsingRecordsetWithTransaction = True
    Exit Function
    
ErrorHandler:
    ' Rollback transaction if there was an error
    
    If Err.Number = 3022 Then
        MsgBox "El producto: " & rsSource!IdProducto & " No se agrega porque ya existe", vbInformation, "Producto ya existe"
        Resume Next
    Else
        ws.Rollback
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in AppendProductosUsingRecordsetWithTransaction"
    
    End If
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
    Set ws = Nothing
    AppendProductosUsingRecordsetWithTransaction = False
End Function

' Optional: Function to check for duplicate product IDs before inserting
Public Function AppendProductosWithDuplicateCheck()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsDestination As DAO.Recordset
    Dim rsCheck As DAO.Recordset
    Dim strSQL As String
    Dim intRecordsAffected As Integer
    Dim intDuplicatesSkipped As Integer
    
    ' Get reference to current database
    Set db = CurrentDb
    
    ' Open source recordset
    Set rsSource = db.OpenRecordset("tblLineaDetalle", dbOpenSnapshot)
    
    ' Open destination recordset
    Set rsDestination = db.OpenRecordset("tblProductos", dbOpenDynaset)
    
    ' Check if source recordset has records
    If rsSource.EOF Then
        MsgBox "No hay registros en tblLineaDetalle.", vbInformation, "Anexar productos"
        GoTo CleanUp
    End If
    
    ' Loop through source recordset
    rsSource.MoveFirst
    intRecordsAffected = 0
    intDuplicatesSkipped = 0
    
    ' Print header for the debug output
    Debug.Print "--- tblLineaDetalle Records ---"
    Debug.Print "IdProducto | Codigo | PorcentajeImpuesto | Detalle | PrecioUnitario | UnidadMedida | UnidadMedidaComercial"
    Debug.Print String(100, "-") ' Separator line
    
    
    Do Until rsSource.EOF
        ' Check if product ID already exists
        strSQL = "SELECT IdProducto FROM tblProductos WHERE IdProducto = '" & Replace(rsSource!IdProducto, "'", "''") & "'"
        Set rsCheck = db.OpenRecordset(strSQL, dbOpenSnapshot)
        
        ' If not a duplicate, add it
        If rsCheck.EOF Then
        
            ' Debug output to Immediate Window (Ctrl+G to view)
            Debug.Print rsSource!IdProducto & " | " & _
                        rsSource!Codigo & " | " & _
                        rsSource!PorcentajeImpuesto & " | " & _
                        Left(rsSource!Detalle, 20) & "... | " & _
                        rsSource!PrecioUnitario & " | " & _
                        rsSource!UnidadMedida & " | " & _
                        rsSource!UnidadMedidaComercial
                    
            ' Add new record to destination
            rsDestination.AddNew
            
            ' Map the fields from source to destination
            rsDestination!IdProducto = rsSource!IdProducto
            rsDestination!CodigoCabys = rsSource!Codigo
            rsDestination!PorcentajeImpuesto = rsSource!PorcentajeImpuesto
            rsDestination!Detalle = rsSource!Detalle
            rsDestination!PrecioUnitario = rsSource!PrecioUnitario
            rsDestination!UnidadMedida = rsSource!UnidadMedida
            rsDestination!UnidadMedidaComercial = rsSource!UnidadMedidaComercial
            rsDestination!Cantidad = 0 ' Hardcoded value as in your SQL
            
            ' Save the new record
            rsDestination.Update
            
            ' Increment counter
            intRecordsAffected = intRecordsAffected + 1
        Else
            ' Increment duplicate counter
            intDuplicatesSkipped = intDuplicatesSkipped + 1
        End If
        
        ' Close check recordset
        rsCheck.Close
        Set rsCheck = Nothing
        
        ' Move to next record
        rsSource.MoveNext
    Loop
    
    ' Report success with duplicate information
    MsgBox intRecordsAffected & " Productos fueron incluidos." & vbCrLf & _
           intDuplicatesSkipped & " duplicados no fueron incluidos.", _
           vbInformation, "Proceso Anexar"
    
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
    AppendProductosWithDuplicateCheck = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in AppendProductosWithDuplicateCheck"
    
    ' Close recordsets if they are open
    If Not rsSource Is Nothing Then
        rsSource.Close
    End If
    If Not rsDestination Is Nothing Then
        rsDestination.Close
    End If
    If Not rsCheck Is Nothing Then
        rsCheck.Close
    End If
    
    Set rsCheck = Nothing
    Set rsSource = Nothing
    Set rsDestination = Nothing
    Set db = Nothing
    AppendProductosWithDuplicateCheck = False
End Function

' Optional: Add a procedure that can be called from a button or other UI element
Public Sub RunAppendProductosUsingRecordset()
    Dim intChoice As Integer
    
    intChoice = MsgBox("Would you like to check for duplicate products before appending?", _
              vbQuestion + vbYesNoCancel, "Append Products")
    
    If intChoice = vbYes Then
        AppendProductosWithDuplicateCheck
    ElseIf intChoice = vbNo Then
        AppendProductosUsingRecordsetWithTransaction
    End If
End Sub
