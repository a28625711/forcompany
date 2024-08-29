Sub ExtractTableInfo()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim emptyTitleRows As Long
    Dim headers As Variant
    Dim currentRow As Variant

    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = "EmptyTitleRows"
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    headers = Application.Transpose(wsSource.Range("A1:E1").Value)
    Set dict = CreateObject("Scripting.Dictionary")
    
    emptyTitleRows = 1
    
    For i = 2 To lastRow
        Set currentRow = CreateObject("Scripting.Dictionary")
        
        Dim j As Integer
        For j = LBound(headers) To UBound(headers)
            currentRow(headers(j)) = wsSource.Cells(i, j + 1).Value
        Next j
        
        dict.Add i - 1, currentRow

        If currentRow("二级标题") = "" Then
            wsDest.Cells(emptyTitleRows, 1).Value = currentRow("编号")
            wsDest.Cells(emptyTitleRows, 2).Value = currentRow("从属编号")
            wsDest.Cells(emptyTitleRows, 3).Value = currentRow("一级标题")
            wsDest.Cells(emptyTitleRows, 4).Value = currentRow("二级标题")
            wsDest.Cells(emptyTitleRows, 5).Value = currentRow("内容")
            emptyTitleRows = emptyTitleRows + 1
        End If
    Next i
    
    MsgBox "数据提取完成！字典中有 " & dict.Count & " 个条目，二级标题为空的行已复制到新工作表。"
End Sub
