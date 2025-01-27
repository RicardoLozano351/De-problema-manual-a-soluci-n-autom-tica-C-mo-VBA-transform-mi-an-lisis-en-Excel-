# De problema manual a solución automática: Cómo VBA transformó mi análisis en Excel 
Sub BuscarBarraYExtraerDatos()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim barra As String
    Dim filaEncabezados As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim colBarra As Long
    Dim nuevaFila As Long
    Dim i As Long

    ' Solicitar al usuario el nombre de la barra a consultar
    barra = Trim(UCase(InputBox("Ingresa el nombre exacto de la barra que deseas consultar:", "Consulta Barra")))
    If barra = "" Then Exit Sub ' Salir si no se ingresa nada

    ' Definir la hoja de origen
    Set wsOrigen = ThisWorkbook.Sheets("Cmg_Barra") ' Cambia el nombre si es necesario
    filaEncabezados = 1 ' Los encabezados están en la fila 1
    lastRow = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row ' Última fila con datos
    lastCol = wsOrigen.Cells(filaEncabezados, wsOrigen.Columns.Count).End(xlToLeft).Column ' Última columna con datos

    ' Buscar la columna de la barra
    colBarra = 0
    For i = 2 To lastCol ' Empezar desde la columna 2 (la primera es fecha/hora)
        If Trim(UCase(wsOrigen.Cells(filaEncabezados, i).Value)) = barra Then
            colBarra = i
            Exit For
        End If
    Next i

    ' Si no se encuentra la barra, mostrar un mensaje
    If colBarra = 0 Then
        MsgBox "No se encontró la barra especificada: " & barra, vbExclamation, "Error"
        Exit Sub
    End If

    ' Crear o seleccionar la hoja de destino
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Datos_" & barra)
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add
        wsDestino.Name = "Datos_" & barra
    End If
    On Error GoTo 0
    wsDestino.Cells.Clear ' Limpiar datos previos

    ' Copiar encabezados a la hoja de destino
    wsDestino.Cells(1, 1).Value = "Fecha/Hora"
    wsDestino.Cells(1, 2).Value = "CMg (" & barra & ")"

    ' Copiar datos de la barra seleccionada
    nuevaFila = 2
    For i = filaEncabezados + 1 To lastRow
        If IsNumeric(wsOrigen.Cells(i, colBarra).Value) Then
            wsDestino.Cells(nuevaFila, 1).Value = wsOrigen.Cells(i, 1).Value ' Fecha/Hora
            wsDestino.Cells(nuevaFila, 2).Value = wsOrigen.Cells(i, colBarra).Value ' Valores
            nuevaFila = nuevaFila + 1
        End If
    Next i

    ' Ajustar formato
    wsDestino.Columns("A:B").AutoFit
    MsgBox "Los datos de la barra '" & barra & "' se han copiado en la hoja '" & wsDestino.Name & "'.", vbInformation, "Consulta completada"
End Sub
