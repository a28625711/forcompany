' TableRow Class
Option Explicit

Public ID As String
Public ParentID As String
Public Level1Title As String
Public Level2Title As String
Public Content As String



Sub ExtractTableData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tableRow As TableRow
    Dim rowsCollection As Collection

    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set rowsCollection = New Collection

    For i = 2 To lastRow
        Set tableRow = New TableRow
        With ws
            tableRow.ID = .Cells(i, 1).Value
            tableRow.ParentID = .Cells(i, 2).Value
            tableRow.Level1Title = .Cells(i, 3).Value
            tableRow.Level2Title = .Cells(i, 4).Value
            tableRow.Content = .Cells(i, 5).Value
        End With
        
        rowsCollection.Add tableRow
    Next i

    For Each tableRow In rowsCollection
        Debug.Print tableRow.ID, tableRow.ParentID, tableRow.Level1Title, tableRow.Level2Title, tableRow.Content
    Next tableRow
End Sub
