Sub AggregateData()
    Dim sendTo As Integer, getFrom As Integer, x As Integer, row As Integer
    For row = 2 To numRows("Input") - 1
        'Starts With Team Number
        getFrom = 5
        sendTo = 1
        sendTo = copy(getFrom, sendTo, row)
        'Yellow Cards
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Auto Scoring
        getFrom = getFrom + 1
        sendTo = gamePieces(getFrom, sendTo, row)
        'Exited Community
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Auto Docking
        getFrom = getFrom + 1
        sendTo = docking(getFrom, sendTo, row, True)
        'Teleop Scoring
        getFrom = getFrom + 1
        sendTo = gamePieces(getFrom, sendTo, row)
        'Fouls
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Tech Fouls
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Red Cards
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Final Status
        getFrom = getFrom + 1
        sendTo = docking(getFrom, sendTo, row, False)
        'Struggled
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Total Docked Bots
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Driver Skill
        getFrom = getFrom + 1
        sendTo = skill(getFrom, sendTo, row)
        'Defense Rating
        getFrom = getFrom + 1
        sendTo = skill(getFrom, sendTo, row)
        'Was Defended
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Died
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'Tippy
        getFrom = getFrom + 1
        sendTo = copy(getFrom, sendTo, row)
        'AutoPoints
        sendTo = AutoPoints("Numerical", row, sendTo)
        'Points
        sendTo = Points("Numerical", row, sendTo)
    Next row
    writeTeams
    For x = 2 To sendTo - 1
        averageColumn (x)
    Next x
End Sub
Function writeTeams()
    Dim row As Integer, rows As Integer, team, checkRow As Integer, switches As Integer, hold As Variant, temp As Variant
    Worksheets("Numerical").Range("A2:A" & (numRows("Numerical") - 1)).copy Worksheets("Average").Range("A2")
    rows = numRows("Average") + 1
    For row = 2 To rows
        team = Worksheets("Average").Range("A" & row).Value
        For checkRow = row + 1 To rows
            If Worksheets("Average").Range("A" & checkRow).Value = team Then
                Worksheets("Average").Range("A" & checkRow).Value = Null
            End If
        Next checkRow
    Next row
End Function
Function Points(sheet As String, row As Integer, sendTo As Integer) As Integer
    Dim val As Double
    val = val + 6 * (Worksheets(sheet).Range("C" & row).Value + Worksheets(sheet).Range("D" & row).Value)
    val = val + 4 * (Worksheets(sheet).Range("E" & row).Value + Worksheets(sheet).Range("F" & row).Value)
    val = val + 3 * Worksheets(sheet).Range("G" & row).Value
    val = val + 3 * Worksheets(sheet).Range("I" & row).Value
    val = val + 8 * Worksheets(sheet).Range("J" & row).Value
    val = val + 5 * (Worksheets(sheet).Range("K" & row).Value + Worksheets(sheet).Range("L" & row).Value)
    val = val + 3 * (Worksheets(sheet).Range("M" & row).Value + Worksheets(sheet).Range("N" & row).Value)
    val = val + 2 * Worksheets(sheet).Range("O" & row).Value
    val = val + 6 * Worksheets(sheet).Range("T" & row).Value
    Worksheets(sheet).Range(columnLetter(sendTo) & row).Value = val
    Points = sendTo + 1
End Function
Function AutoPoints(sheet As String, row As Integer, sendTo As Integer) As Integer
    Dim val As Double
    val = val + 6 * (Worksheets(sheet).Range("C" & row).Value + Worksheets(sheet).Range("D" & row).Value)
    val = val + 4 * (Worksheets(sheet).Range("E" & row).Value + Worksheets(sheet).Range("F" & row).Value)
    val = val + 3 * Worksheets(sheet).Range("G" & row).Value
    val = val + 3 * Worksheets(sheet).Range("I" & row).Value
    val = val + 8 * Worksheets(sheet).Range("J" & row).Value
    Worksheets(sheet).Range(columnLetter(sendTo) & row).Value = val
    AutoPoints = sendTo + 1
End Function
Function averageColumn(column As Integer)
    Dim row As Integer, val As Double, div As Integer, team, x
    For row = 2 To numRows("Average") - 1
        val = 0
        div = 0
        team = Worksheets("Average").Range("A" & row).Value
        For x = 2 To numRows("Numerical")
            If Worksheets("Numerical").Range("A" & x).Value = team Then
                If Not Worksheets("Numerical").Range(columnLetter(column) & x).Value < 0 Then
                    val = val + Worksheets("Numerical").Range(columnLetter(column) & x).Value
                    div = div + 1
                End If
            End If
        Next x
        If div = 0 Then
            Worksheets("Average").Range(columnLetter(column) & row).Value = 0
        Else
            Worksheets("Average").Range(columnLetter(column) & row).Value = val / div
        End If
    Next row
End Function
Function numRows(Worksheet As String) As Integer
    Dim repeat As Boolean
    repeat = True
    numRows = 1
    Do While repeat
        If IsEmpty(Worksheets(Worksheet).Range("A" & numRows)) Then
            repeat = False
        Else
        numRows = numRows + 1
        End If
    Loop
End Function
Function gamePieces(getFrom As Integer, sendTo As Integer, row As Integer) As Integer
    Dim hiCo As Integer, hiCu As Integer, miCo As Integer, miCu As Integer, loPi As Integer, numPieces As Integer, pieces As Variant, piece As Variant, modNumber As Integer, cube As Boolean, high As Boolean, low As Boolean, mid As Boolean
    pieces = Split(Worksheets("Input").Range(columnLetter(getFrom) & row).Value, ",")
    numPieces = ArrayLen(pieces)
    For Each piece In pieces
        modNumber = (piece + 1) Mod 3
        cube = (modNumber = 0)
        If piece < 10 Then
            If cube Then
                hiCu = hiCu + 1
            Else
                hiCo = hiCo + 1
            End If
        Else
            If piece > 18 Then
                loPi = loPi + 1
            Else
                If cube Then
                    miCu = miCu + 1
                Else
                    miCo = miCo + 1
                End If
            End If
        End If
    Next piece
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = hiCo
    sendTo = sendTo + 1
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = hiCu
    sendTo = sendTo + 1
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = miCo
    sendTo = sendTo + 1
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = miCu
    sendTo = sendTo + 1
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = loPi
    sendTo = sendTo + 1
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = numPieces
    sendTo = sendTo + 1
    gamePieces = sendTo
End Function
Function copy(getFrom As Integer, sendTo As Integer, row As Integer) As Integer
    Dim val As Variant
    val = Worksheets("Input").Range(columnLetter(getFrom) & row).Value
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = val
    copy = sendTo + 1
End Function
Function docking(getFrom As Integer, sendTo As Integer, row As Integer, auto As Boolean) As Integer
    Dim Value As Variant
    Value = Worksheets("Input").Range(columnLetter(getFrom) & row).Value
    Select Case (Value)
        Case "p":
            Value = 1 / 3
        Case "e":
            If auto Then
                Value = 1.5
            Else
                Value = 5 / 3
            End If
        Case "d":
            Value = 1
        Case "x":
            Value = -1
        Case "a":
            Value = 0
    End Select
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = Value
    docking = sendTo + 1
End Function
Function skill(getFrom As Integer, sendTo As Integer, row As Integer) As Integer
    Dim val As Variant, x As Double
    val = Worksheets("Input").Range(columnLetter(getFrom) & row).Value
    Select Case (val)
        Case "x":
            x = -1
        Case "b":
            x = 0
        Case "a":
            x = 1
        Case "aa":
            x = 2
    End Select
    Worksheets("Numerical").Range(columnLetter(sendTo) & row).Value = x
    skill = sendTo + 1
End Function
Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
Function columnLetter(columnNumber As Integer) As String
    columnLetter = Split(Cells(1, columnNumber).Address, "$")(1)
End Function
Sub sortByColumn(column As String, Worksheet As String)
    Dim row As Integer, hold1 As Variant, hold2 As Variant, switches As Integer
    switches = 1
    Do While Not switches = 0
        switches = 0
        For row = 2 To numRows(Worksheet) - 2
            If Worksheets(Worksheet).Range(column & row).Value < Worksheets(Worksheet).Range(column & (row + 1)).Value Then
                hold1 = Worksheets(Worksheet).Range("A" & row & ":Z" & row)
                hold2 = Worksheets(Worksheet).Range("A" & (row + 1) & ":AZ" & (row + 1))
                Worksheets(Worksheet).Range("A" & row & ":Z" & row) = hold2
                Worksheets(Worksheet).Range("A" & (row + 1) & ":Z" & (row + 1)) = hold1
                switches = switches + 1
            End If
        Next row
    Loop
End Sub
Sub sortByColumnInverse(column As String, Worksheet As String)
    Dim row As Integer, hold1 As Variant, hold2 As Variant, switches As Integer
    switches = 1
    Do While Not switches = 0
        switches = 0
        For row = 2 To numRows(Worksheet) - 2
            If Worksheets(Worksheet).Range(column & row).Value > Worksheets(Worksheet).Range(column & (row + 1)).Value Then
                hold1 = Worksheets(Worksheet).Range("A" & row & ":Z" & row)
                hold2 = Worksheets(Worksheet).Range("A" & (row + 1) & ":AZ" & (row + 1))
                Worksheets(Worksheet).Range("A" & row & ":Z" & row) = hold2
                Worksheets(Worksheet).Range("A" & (row + 1) & ":Z" & (row + 1)) = hold1
                switches = switches + 1
            End If
        Next row
    Loop
End Sub
Sub prcss1QRCodeInput()
    saveData (getInput())
End Sub
Sub prcss6QRCodeInput()
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
End Sub
Public Function getInput()
    getInput = InputBox("Scan QR Code", "Match Scouting Input")
End Function
Sub testSaveData()
    saveData ("s=fff;e=1234;l=qm;m=1234;r=r1;t=1234;as=;ae=Y;al=2;ao=2;ai=1;aa=Y;at=N;ax=Y;lp=2;op=1;ip=3;rc=pass;f=0;pc=pass;ss=;c=pass;b=N;ca=x;cb=x;cs=slow;p=N;ds=x;dr=x;pl=x;tr=N;wd=N;if=N;d=N;to=N;be=N;cf=N")
End Sub
Sub saveData(inp As String)
    Dim fields
    Dim par() As String
    Dim Value
    Dim Key
    Dim table As ListObject
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim mapper
    Set mapper = CreateObject("Scripting.Dictionary")
    Dim data
    Set data = CreateObject("Scripting.Dictionary")
    Dim tableName As String
    tableName = "ScoutingData"
    ' Set up map
    ' Fields for every year
    mapper.Add "s", "scouter"
    mapper.Add "e", "eventCode"
    mapper.Add "l", "matchLevel"
    mapper.Add "m", "matchNumber"
    mapper.Add "r", "robot"
    mapper.Add "t", "teamNumber"
    mapper.Add "as", "autoStartPosition"
    mapper.Add "asg", "autoScoring"
    mapper.Add "ec", "exitedCommunity"
    mapper.Add "ad", "autoDocking"
    mapper.Add "agpa", "autoAttemptedPieces"
    mapper.Add "gph", "gamePiecesStillWithBot"
    mapper.Add "tct", "Cycles"
    mapper.Add "tsg", "teleopScoring"
    mapper.Add "dt", "dockingTimer"
    mapper.Add "fs", "finalStatus"
    mapper.Add "stg", "struggled"
    mapper.Add "dn", "totalDockedBots"
    mapper.Add "ds", "driverSkill"
    mapper.Add "dr", "defenseRating"
    mapper.Add "wd", "wasDefended"
    mapper.Add "die", "died/immobilized"
    mapper.Add "fl", "fouls"
    mapper.Add "tf", "techFouls"
    mapper.Add "yc", "yellow"
    mapper.Add "rc", "red"
    mapper.Add "tip", "Tippy?"
    mapper.Add "co", "Comments"
    ' Additional custom mapping
    'mapper.Add "f", "fouls"
    'mapper.Add "c", "climb"
    'mapper.Add "dr", "defenseRating"
    'mapper.Add "d", "died"
    'mapper.Add "to", "tippedOver"
    'mapper.Add "cf", "cardFouls"
    'mapper.Add "co", "comments"
    If inp = "" Then
        Exit Sub
    End If
    'MsgBox (inp)
    fields = Split(inp, ";")
    If ArrayLen(fields) > 0 Then
        Dim i As Integer
        Dim str
        i = 0
        For Each str In fields
            par = Split(str, "=")
            Key = par(0)
            Value = par(1)
            If mapper.Exists(Key) Then
                Key = mapper(Key)
            End If
            data.Add Key, Value
        Next
        tableexists = False
        Dim tbl As ListObject
        Dim sht As Worksheet
        'Loop through each sheet and table in the workbook
        For Each sht In ThisWorkbook.Worksheets
            For Each tbl In sht.ListObjects
                If tbl.Name = tableName Then
                    tableexists = True
                    Set table = tbl
                    Set ws = sht
                End If
            Next tbl
        Next sht
        If tableexists Then
            'Set table = ws.ListObjects(tableName)
        Else
            Dim tablerange As Range
            ws.ListObjects.Add(xlSrcRange, Range("A1:AO1"), , xlYes).Name = tableName
            i = 0
            Set table = ws.ListObjects(tableName)
            For Each Key In data.Keys
                table.Range(i + 1) = Key
                i = i + 1
            Next
        End If
        Dim newrow As ListRow
        Set newrow = table.ListRows.Add
        For Each str In data.Keys
            newrow.Range(table.ListColumns(str).Index) = data(str)
        Next
    End If
End Sub
