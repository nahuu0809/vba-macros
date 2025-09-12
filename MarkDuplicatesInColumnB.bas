Attribute VB_Name = "Module1"
Sub MarkDuplicatesInColumnB()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim email As String
    Dim duplicateCount As Integer
    
    ' Establecer la hoja activa
    Set ws = ActiveSheet
    
    ' Encontrar la última fila y columna con datos
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Definir el rango de la columna B (desde B1 hasta la última fila con datos)
    Set rng = ws.Range("B1:B" & lastRow)
    
    ' Crear un diccionario para almacenar correos y contar duplicados
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Primera pasada: identificar duplicados en la columna B (ignorando mayúsculas)
    For Each cell In rng
        email = Trim(LCase(cell.Value)) ' Convertir a minúsculas y limpiar espacios
        If email <> "" Then ' Ignorar celdas vacías
            If dict.exists(email) Then
                dict(email) = dict(email) + 1
            Else
                dict.Add email, 1
            End If
        End If
    Next cell
    
    ' Segunda pasada: marcar filas (verde para no duplicados, rojo para duplicados)
    For Each cell In rng
        email = Trim(LCase(cell.Value)) ' Convertir a minúsculas para comparar
        If email <> "" Then
            ' Seleccionar toda la fila hasta la última columna con datos
            If dict(email) > 1 Then
                ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, lastCol)).Interior.ColorIndex = 3 ' Rojo para duplicados
            Else
                ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, lastCol)).Interior.ColorIndex = 4 ' Verde para no duplicados
            End If
        End If
    Next cell
    
    ' Contar el total de filas duplicadas
    duplicateCount = 0
    For Each Key In dict.keys
        If dict(Key) > 1 Then
            duplicateCount = duplicateCount + dict(Key)
        End If
    Next Key
    
    MsgBox "Se han marcado " & duplicateCount & " filas con correos duplicados en rojo y " & (lastRow - duplicateCount) & " filas con correos no duplicados en verde.", vbInformation
End Sub
