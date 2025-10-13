Dim ws As Worksheet
    Dim editRow As Long
    Dim aoColumn As Long
    Dim cellAddress As String
    Dim lastRow As Long
    Dim nextRow As Long
    
    Set ws = ActiveSheet
    aoColumn = ws.Range("AO1").Column
    editRow = Target.Row
    
    If Target.Column = aoColumn Then Exit Sub
    If editRow = 1 Then Exit Sub
    
    cellAddress = Target.Address(False, False)
    Application.EnableEvents = False
    Target.Interior.Color = RGB(255, 255, 200)
    
    lastRow = ws.Cells(ws.Rows.Count, aoColumn).End(xlUp).Row
    
    If ws.Cells(1, aoColumn).Value = "" Then
        nextRow = 1
    Else
        nextRow = lastRow + 1
    End If
    
    ws.Cells(nextRow, aoColumn).Value = cellAddress
    Application.EnableEvents = True
End Sub

Sub ToggleColoring()
    coloringEnabled = Not coloringEnabled
    
    If coloringEnabled Then
        MsgBox "確定しました（色付け機能ON）"
    Else
        MsgBox "確定を解除しました（色付け機能OFF）"
    End If
    
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets(1)
    targetSheet.Range("AP1").Value = coloringEnabled
End Sub
