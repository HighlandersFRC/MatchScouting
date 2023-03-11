Function ColumnLetter(columnNumber As Long) As String
    ColumnLetter = Split(Cells(1, columnNumber).Address, "$")(1)
End Function
Function gamePieces(getFrom As Integer, sendTo As Integer, row As Integer)
    Dim hiCo As Integer, hiCu As Integer, miCo As Integer, miCu As Integer, loCo As Integer, loCu As Integer
    Split(Worksheets("ScoutingPASS_Excel_Example").Range(row & getFrom).Value, ",")
End Function