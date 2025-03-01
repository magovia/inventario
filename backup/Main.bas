Attribute VB_Name = "Main"
Option Compare Database

Sub extraerFactura()

On Error GoTo ErrorHandler
DoCmd.SetWarnings False


'Borrar datos de la tabla LineaDetalles
    Append.RunActionQuery "BorrarTodoTblLineaDetalle", "Borrar tblLineaDetalle"

'Rellenar tabla lineaDetalle con nuevos valores de la factura cargada
    'Append.RellenarLineaDetalleFromXLS
    Append.RunActionQuery "RellenarTblLineaDetalle", "Rellenar tblLineaDetalle"

'Actualizar campo IdProduct con Ids no repetidos
    uniqueIds.MakeUniqueIds
        
'MsgBox "extraerFactura - done", vbDefaultButton1, "Extraidas"

DoCmd.SetWarnings True

Exit Sub

ErrorHandler:
        MsgBox "Descripcion: " & Err.Description, , "RellenarTblLineaDetalle:Error No." & Err.Number
        DoCmd.SetWarnings True
    Exit Sub
 
End Sub

Sub guardarFactura()

On Error GoTo ErrorHandler
DoCmd.SetWarnings False


'Agregar Proveedores
    Append.RunActionQuery "RellenarTblProveedor", "Guardar factura: RellenarTblProveedor"

'Agregar la factura
    Append.RunActionQuery "RellenarTblFacturas", "Guardar factura: RellenarTblFacturas"
        
 'Agrega los productos desde la tabla LineaDetalle hacia TblProductos
    'Append.RunActionQuery "RellenarTblProductos", "Guardar factura: RellenarTblProductos"
    appendProductos.AppendProductosWithDuplicateCheck
      
 'Suma las cantidades en la tabla de Productos desde la tblLineaDetalles
     Append.RunActionQuery "AgregarInventario", "Guardar factura: AgregarInventario"

'Agregar lineas de factura a tblFacturasDetalle
    'Append.RunActionQuery "RellenarTblFacturasDetalle", "Guardar factura: RellenarTblFacturasDetalle"
    appFacturaDetalle.AppendFacturasDetalleUsingRecordset
    
'Actualiza los precios unitarios desde lineaDetalle hacia tblProductos
    UpdtPrecios.Precio

'Borrar datos de la tabla LineaDetalles
    Append.RunActionQuery "BorrarTodoTblLineaDetalle", "Guardar factura: BorrarTodoTblLineaDetalle"
  
Debug.Print "guardarFactura - done"

'MsgBox "Factura ha sido guardada exitosamente", vbDefaultButton1, "Registro"

    DoCmd.Close
    
DoCmd.SetWarnings True

Exit Sub

ErrorHandler:
        Debug.Print "guardarFactura: " & Err.Description & " " & Err.Number
        DoCmd.SetWarnings True
    Exit Sub
        
        
End Sub




