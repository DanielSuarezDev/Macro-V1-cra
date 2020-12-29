Sub ConsultarPagos()
Dim Sql As String
Dim Criterio  As String
Dim CriterioFechaInicial, CriterioFechaFinal As String
Dim limpiarDatos, lngCampos, i As Long
    limpiarDatos = Sheets("PAGOS").Range("A" & Rows.Count).End(xlUp).Row
    Sheets("PAGOS").Range("A2:D" & limpiarDatos).ClearContents
    Call ConectarBase
    Set rsConciliacion = New ADODB.Recordset
    
    If rsConciliacion.State = 1 Then
        rsConciliacion.Close
    End If
    Criterio = Trim(Sheets("CONTABILIZADOS").Cells(1, 5))
    CriterioFechaInicial = Mid(Sheets("CONTABILIZADOS").Cells(2, 5), 4, 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(2, 5), 2) & "/" & Right(Sheets("CONTABILIZADOS").Cells(2, 5), 4)
    CriterioFechaFinal = Mid(Sheets("CONTABILIZADOS").Cells(3, 5), 4, 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(3, 5), 2) & "/" & Right(Sheets("CONTABILIZADOS").Cells(3, 5), 4)
    Sql = "SELECT * FROM Conciliacion WHERE  [Fecha de Transmisión] BETWEEN #" & CriterioFechaInicial & "# AND  #" & CriterioFechaFinal & "# "
    rsConciliacion.Open Sql, Conexion
    Sheets("Prueba").Cells(2, 1).CopyFromRecordset rsConciliacion
    lngCampos = rsConciliacion.Fields.Count
    For i = 0 To lngCampos - 1
         Sheets("PAGOS").Cells(1, i + 1).Value = rsConciliacion.Fields(i).Name
    Next
    limpiarDatos = Sheets("PAGOS").Range("A" & Rows.Count).End(xlUp).Row
    For i = 2 To limpiarDatos
                Sheets("PAGOS").Cells(i, 1).Value = CLng(Sheets("PAGOS").Cells(i, 1))
    Next i
    rsConciliacion.Close
    Set rsConciliacion = Nothing
    Conexion.Close
    Set Conexion = Nothing
End Sub

