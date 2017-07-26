Sub Test1()

    ' Test1 Macro
    ' This macro finds the columns with the known headings and changes their formats if found


    ' 01. Find the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet


    ' 02. Find the cells on the header row
    Dim headerRowNumber As Integer: headerRowNumber = 1 ' Configurable - just in case the header is not on row 1
    Dim startCol As String: startCol = "A" ' Configurable - just in case the first column is not A
    Dim endCol As String: endCol = "Z" ' Configurable - just in case the last column is not Z
    Dim startCellName As String: startCellName = startCol & headerRowNumber
    Dim endCellName As String: endCellName = endCol & headerRowNumber

    Dim cellRange As Range
    Set cellRange = ws.Cells.Range(startCellName, endCellName)


    ' 03. List all known column names in an array
    Dim knownCols(5) As String

    knownCols(0) = "Header1"
    knownCols(1) = "Header3"
    knownCols(2) = "Header5"
    knownCols(3) = "Header7"
    knownCols(4) = "Header9"


    ' 04. Change the data type of all known columns
    Dim cell As Range
    Dim colNum As Integer

    For Each knownCol In knownCols
        Set cell = cellRange.Find(What:=knownCol, LookAt:=xlWhole, MatchCase:=False)
        If cell Is Nothing Then
            MsgBox "Could not find column with header = " & knownCol
        Else
            colNum = cell.Column
            ws.Columns(colNum).NumberFormat = "#,##0.00"
        End If
    Next knownCol

    MsgBox "All known columns have been formatted"

End Sub
