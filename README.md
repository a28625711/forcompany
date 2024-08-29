# forcompany
Private

https://1drv.ms/i/c/7f34c6d1925ecc82/EZMk2qtwFEVFhNb8e3xEXH0BPkN5ClzP0JGyDR0RKbsAnQ

https://1drv.ms/i/c/7f34c6d1925ecc82/EecYfx5a9vZEnYISWN0a3UIB9uqfhQxCcwtHap8WAYqBxg


Sub ExtractTableInfo()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim key As String
    Dim emptyTitleRows As Long

    ' 设置源工作表
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' 修改为你的源工作表名称
    ' 创建一个新的工作表用于存储二级标题为空的行
    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = "EmptyTitleRows"
    
    ' 获取数据最后一行
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    ' 创建字典对象
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 初始化空标题行计数
    emptyTitleRows = 1
    
    ' 从第二行开始遍历
    For i = 2 To lastRow
        key = wsSource.Cells(i, 1).Value ' 编号
        
        ' 存入字典
        If Not dict.Exists(key) Then
            dict.Add key, Array(wsSource.Cells(i, 2).Value, wsSource.Cells(i, 3).Value, wsSource.Cells(i, 4).Value, wsSource.Cells(i, 5).Value)
        End If
        
        ' 检查二级标题是否为空
        If wsSource.Cells(i, 4).Value = "" Then
            wsDest.Cells(emptyTitleRows, 1).Value = wsSource.Cells(i, 1).Value ' 编号
            wsDest.Cells(emptyTitleRows, 2).Value = wsSource.Cells(i, 2).Value ' 从属编号
            wsDest.Cells(emptyTitleRows, 3).Value = wsSource.Cells(i, 3).Value ' 一级标题
            wsDest.Cells(emptyTitleRows, 4).Value = wsSource.Cells(i, 4).Value ' 二级标题
            wsDest.Cells(emptyTitleRows, 5).Value = wsSource.Cells(i, 5).Value ' 内容
            emptyTitleRows = emptyTitleRows + 1
        End If
    Next i
    
    ' 提示信息
    MsgBox "数据提取完成！字典中有 " & dict.Count & " 个条目，二级标题为空的行已复制到新工作表。"
End Sub
