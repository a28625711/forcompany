Sub ExtractTableInfo()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim emptyTitleRows As Long
    Dim headers As Variant
    Dim currentRow As Variant

    ' 设置源工作表
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' 修改为你的源工作表名称
    ' 创建一个新的工作表用于存储二级标题为空的行
    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = "EmptyTitleRows"
    
    ' 获取数据最后一行
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    ' 读取标题行
    headers = Application.Transpose(wsSource.Range("A1:E1").Value)
    
    ' 创建字典对象
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 初始化空标题行计数
    emptyTitleRows = 1
    
    ' 从第二行开始遍历
    For i = 2 To lastRow
        ' 创建当前行数据的字典
        Set currentRow = CreateObject("Scripting.Dictionary")
        
        ' 遍历每一列并存入字典
        Dim j As Integer
        For j = LBound(headers) To UBound(headers)
            currentRow(headers(j)) = wsSource.Cells(i, j + 1).Value
        Next j
        
        ' 将当前行字典添加到主字典
        dict.Add i - 1, currentRow ' 使用行索引作为主字典的键

        ' 检查二级标题是否为空
        If currentRow("二级标题") = "" Then
            wsDest.Cells(emptyTitleRows, 1).Value = currentRow("编号")
            wsDest.Cells(emptyTitleRows, 2).Value = currentRow("从属编号")
            wsDest.Cells(emptyTitleRows, 3).Value = currentRow("一级标题")
            wsDest.Cells(emptyTitleRows, 4).Value = currentRow("二级标题")
            wsDest.Cells(emptyTitleRows, 5).Value = currentRow("内容")
            emptyTitleRows = emptyTitleRows + 1
        End If
    Next i
    
    ' 提示信息
    MsgBox "数据提取完成！字典中有 " & dict.Count & " 个条目，二级标题为空的行已复制到新工作表。"
End Sub
