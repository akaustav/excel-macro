Sub Test1()
'
' Test1 Macro
' This macro finds the column with the heading Ameet and displays a message box containing the cell number
'

' Loop thru all column headings and find the cell with the column heading matching "Ameet"
' Convert the column type to currency

    ' 01. Find the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    
    ' 02. Find the cells on the header row
    Dim headerRowNumber As Integer
    Dim startCol As String
    Dim endCol As String
    Dim startCellName As String
    Dim endCellName As String
    Dim cellRange As Range
    Dim cellValue As String

    headerRowNumber = 1 ' Configurable - just in case the header is not on row 1
    startCol = "A" ' Configurable - just in case the first column is not A
    endCol = "Z" ' Configurable - just in case the last column is not Z
    startCellName = startCol & headerRowNumber
    endCellName = endCol & headerRowNumber
    Set cellRange = ws.Cells.Range(startCellName, endCellName)


    ' 03. Change the data type of the columns
    Dim message As String
    message = "Values in cells from " & startCellName & " thru " & endCellName & " are:"

    Dim cell As Range
    Dim colNum As Integer

    For Each cell In cellRange
        If Trim(cell.Value) <> "" Then
            colNum = cell.Column
            ws.Columns(colNum).NumberFormat = "#,##0.00"
            message = message & vbNewLine & "Col" & cell.Column & " = " & cell.Value
        End If
    Next cell

    MsgBox message

'    Dim rng As Range
'    Dim colHeading As String

'    colHeading = "Ameet"

'    With Worksheets(activeWorksheet).Range("A1:BB1")
'        Set rng = Worksheets(activeWorksheet).Range("A1:BB1").Find(What:=colHeading, LookAt:=xlWhole, MatchCase:=False)
'        Do While Not rng Is Nothing
'            MsgBox "The column with " & colHeading & " is at " & rng
'            Set rng = .FindNext
'        Loop
'    End With

End Sub