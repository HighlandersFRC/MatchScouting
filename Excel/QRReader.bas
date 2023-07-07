Sub AggregateData()
    Dim sendto As Integer, getFrom As Integer, x As Integer, row As Integer, foulSpot As Integer, lastColumn As Integer, struggled As Integer, cards As Integer
    For row = 2 To numRows("Input") - 1
        sendto = 0
        'Starts With Match Number
        getFrom = 3
        sendto = sendto + 1
        sendto = copy(getFrom, sendto, row)
        'Team Number
        getFrom = getFrom + 2
        sendto = copy(getFrom, sendto, row)
        'skip three for points
        sendto = sendto + 3
        'StartPosition
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'autoHighCones
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'autoHighCubes
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'autoMidCones
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'autoMidCubes
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'autoLowPieces
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'piecesMissed
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Exited Community
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Auto Docking
        getFrom = getFrom + 1
        sendto = docking(getFrom, sendto, row, True)
        'Auto Strategy
        getFrom = getFrom + 1
        'highCones
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'highCubes
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'midCones
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'midCubes
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'lowPieces
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Saving the position from which to grab fouls
        foulSpot = getFrom + 1
        getFrom = getFrom + 4
        'Teleop Strategy
        getFrom = getFrom + 1
        'Final Status
        getFrom = getFrom + 1
        sendto = docking(getFrom, sendto, row, False)
        'Struggled
        getFrom = getFrom + 1
        struggled = sendto
        sendto = copy(getFrom, sendto, row)
        'Total Docked Bots
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Endgame Strategy
        getFrom = getFrom + 1
        'Driver Skill
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Defense Rating
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Died
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Tippy
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'allianceScore
        getFrom = getFrom + 1
        Dim alliancePts As Integer
        alliancePts = sendto
        sendto = copy(getFrom, sendto, row)
        'oppositionAllianceScore
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Fouls
        getFrom = foulSpot
        sendto = copy(getFrom, sendto, row)
        'Tech Fouls
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        'Yellow Cards
        getFrom = getFrom + 1
        cards = sendto
        sendto = copy(getFrom, sendto, row)
        'Red Cards
        getFrom = getFrom + 1
        sendto = copy(getFrom, sendto, row)
        lastColumn = sendto - 1
        'Auto Points
        sendto = autoPoints(3, row)
        'Points
        sendto = points(4, row)
        'contribution
        sendto = contribution(4, alliancePts, 5, row)
    Next row
    writeTeams
    For x = 3 To lastColumn
        Select Case (x):
            Case yc:
                hasCard (x)
            Case rc:
                hasCard (x)
            Case Else:
                getFrom = averageColumn(x)
        End Select
    Next x
End Sub
Sub checkErrors()
    highlightEntries
    duplicateStations
    checkNumEntries
End Sub
Function autoPoints(sendto As Integer, row As Integer) As Integer
    Dim ahc As Integer, ahcu As Integer, amc As Integer, amcu As Integer, alp As Integer, ec As Integer, ad As Integer, thc As Integer, thcu As Integer, tmc As Integer, tmcu As Integer, tlp As Integer, td As Integer, yc As Integer, rc As Integer
    Dim ahcval, ahcuval, amcval, amcuval, alpval, ecval, adval, sheet As String, val As Double
    sheet = "Numerical"
    ahc = 7
    ahcu = 8
    amc = 9
    amcu = 10
    alp = 11
    ec = 13
    ad = 14
    thc = 15
    thcu = 16
    tmc = 17
    tmcu = 18
    tlp = 19
    td = 20
    yc = 30
    rc = 31
    ahcuval = Worksheets(sheet).Range(columnLetter(ahcu) & row).Value
    ahcval = Worksheets(sheet).Range(columnLetter(ahc) & row).Value
    amcval = Worksheets(sheet).Range(columnLetter(amc) & row).Value
    amcuval = Worksheets(sheet).Range(columnLetter(amcu) & row).Value
    alpval = Worksheets(sheet).Range(columnLetter(alp) & row).Value
    ecval = Worksheets(sheet).Range(columnLetter(ec) & row).Value
    adval = Worksheets(sheet).Range(columnLetter(ad) & row).Value
    If adval < 0 Then
        adval = 0
    End If
    val = 0
    val = val + (ahcuval + ahcval) * 6
    val = val + (amcval + amcuval) * 4
    val = val + (ecval + alpval) * 3
    val = val + adval
    Worksheets(sheet).Range(columnLetter(sendto) & row).Value = val
    autoPoints = sendto + 1
End Function
Function points(sendto As Integer, row As Integer) As Integer
    Dim ahc As Integer, ahcu As Integer, amc As Integer, amcu As Integer, alp As Integer, ec As Integer, ad As Integer, thc As Integer, thcu As Integer, tmc As Integer, tmcu As Integer, tlp As Integer, td As Integer, yc As Integer, rc As Integer
    Dim ahcval, ahcuval, amcval, amcuval, alpval, ecval, adval, thcval, thcuval, tmcval, tmcuval, tlpval, tdval, sheet As String, val As Double
    sheet = "Numerical"
    ahc = 7
    ahcu = 8
    amc = 9
    amcu = 10
    alp = 11
    ec = 13
    ad = 14
    thc = 15
    thcu = 16
    tmc = 17
    tmcu = 18
    tlp = 19
    td = 20
    yc = 30
    rc = 31
    ahcuval = Worksheets(sheet).Range(columnLetter(ahcu) & row).Value
    ahcval = Worksheets(sheet).Range(columnLetter(ahc) & row).Value
    amcval = Worksheets(sheet).Range(columnLetter(amc) & row).Value
    amcuval = Worksheets(sheet).Range(columnLetter(amcu) & row).Value
    alpval = Worksheets(sheet).Range(columnLetter(alp) & row).Value
    ecval = Worksheets(sheet).Range(columnLetter(ec) & row).Value
    adval = Worksheets(sheet).Range(columnLetter(ad) & row).Value
    thcuval = Worksheets(sheet).Range(columnLetter(thcu) & row).Value
    thcval = Worksheets(sheet).Range(columnLetter(thc) & row).Value
    tmcval = Worksheets(sheet).Range(columnLetter(tmc) & row).Value
    tmcuval = Worksheets(sheet).Range(columnLetter(tmcu) & row).Value
    tlpval = Worksheets(sheet).Range(columnLetter(tlp) & row).Value
    tdval = Worksheets(sheet).Range(columnLetter(td) & row).Value
    If tdval < 0 Then
        tdval = 0
    End If
    If adval < 0 Then
        adval = 0
    End If
    val = 0
    val = val + (ahcuval + ahcval) * 6
    val = val + (thcuval + thcval) * 5
    val = val + (amcval + amcuval) * 4
    val = val + (ecval + alpval + tmcval + tmcuval) * 3
    val = val + 2 * tlpval
    val = val + adval + tdval
    Worksheets(sheet).Range(columnLetter(sendto) & row).Value = val
    points = sendto + 1
End Function
Function docking(getFrom As Integer, sendto As Integer, row As Integer, auto As Boolean) As Integer
    Dim Value As Variant
    Value = Worksheets("Input").Range(columnLetter(getFrom) & row).Value
    Select Case (Value)
        Case "p":
            Value = 2
        Case "e":
            If auto Then
                Value = 12
            Else
                Value = 10
            End If
        Case "d":
            If auto Then
                Value = 8
            Else
                Value = 6
            End If
        Case "x":
            Value = -1
        Case "a":
            Value = 0
    End Select
    Worksheets("Numerical").Range(columnLetter(sendto) & row).Value = Value
    docking = sendto + 1
End Function
Function sumColumn(column As Integer)
    Dim row As Integer, val As Double
    For row = 2 To numRows("Average") - 1
        val = 0
        team = Worksheets("Average").Range("A" & row).Value
        For x = 2 To numRows("Numerical")
            If Worksheets("Numerical").Range("B" & x).Value = team Then
                If Not Worksheets("Numerical").Range(columnLetter(column) & x).Value < 0 Then
                        val = val + Worksheets("Numerical").Range(columnLetter(column) & x).Value
                End If
            End If
        Next x
        Worksheets("Average").Range(columnLetter(column - 1) & row).Value = val
    Next row
End Function
Sub syncPit()
    Dim rows As Integer, teamRow, hasTeam As Boolean, team, rng
    For rows = 2 To numRows("PitScouting")
        hasTeam = False
        team = Worksheets("PitScouting").Range("B" & rows).Value
        For teamRow = 2 To numRows("Average")
            If Worksheets("Average").Range("A" & teamRow).Value = team Then
                hasTeam = True
                Exit For
            End If
        Next teamRow
    If hasTeam Then
        rng = Sheets("PitScouting").Range("A" & rows & ":R" & rows)
        Sheets("Average").Range("AD" & teamRow) = rng
    End If
    Next rows
End Sub
Sub writeLinks()
    Dim links As Double, team As Integer, eventName As String, row As Integer, json As Object, x As Integer
    eventName = InputBox("What is the event key?(don't include a year)")
    For row = 2 To numRows("Average") - 1
        team = Worksheets("Average").Range("A" & row).Value
        links = GetPolarForecastData(eventName, team)("linkPoints")
        Worksheets("Average").Range("AD" & row).Value = links
        Worksheets("Average").Range("C" & row).Value = Worksheets("Average").Range("C" & row).Value + links
    Next row
End Sub
Sub oprCombos()
    Dim tableexists As Boolean, table As ListObject, ws As Worksheet, tablename As String, teams As Object, eventKey As String, team As Variant, teamnum As Integer, i As Integer, row As ListRow, polar As Object, stat As Object, tba As Object
    eventKey = InputBox("eventKey")
    Set ws = ActiveSheet
    tablename = "OPRs"
    tableexists = False
        Dim tbl As ListObject
        Dim sht As Worksheet
        'Loop through each sheet and table in the workbook
        For Each sht In ThisWorkbook.Worksheets
            For Each tbl In sht.ListObjects
                If tbl.Name = tablename Then
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
            ws.ListObjects.Add(xlSrcRange, Range("A1:H1"), , xlYes).Name = tablename
            i = 0
            Set table = ws.ListObjects(tablename)
            table.Range(1, i + 1) = "Team"
            i = i + 1
            table.Range(1, i + 1) = "TBA + Polar + Statbotics"
            i = i + 1
            table.Range(1, i + 1) = "Polar + Statbotics"
            i = i + 1
            table.Range(1, i + 1) = "TBA + Statbotics"
            i = i + 1
            table.Range(1, i + 1) = "TBA + Polar"
            i = i + 1
            table.Range(1, i + 1) = "Statbotics"
            i = i + 1
            table.Range(1, i + 1) = "TBA"
            i = i + 1
            table.Range(1, i + 1) = "Polar"
            i = i + 1
        End If
        Set teams = GetTeams(eventKey)
        Set tba = GetOPRs(eventKey)
        i = 0
        For i = table.ListRows.Count To 1 Step -1
                table.ListRows(i).Delete
        Next i
        i = 0
        For Each team In teams
            i = i + 1
            table.ListRows.Add.Range(table.ListColumns("Team").Index).Value = team("team_number")
            table.ListRows(i).Range(table.ListColumns("TBA").Index).Value = tba("oprs")(team("key"))
            Set polar = GetPolarForecastData((Split(eventKey, "2023")(1)), (team("team_number")))
            table.ListRows(i).Range(table.ListColumns("Polar").Index).Value = polar("OPR")
            Set stat = GetStatboticsData(eventKey, (team("team_number")))
            table.ListRows(i).Range(table.ListColumns("Statbotics").Index).Value = stat("epa_end")
            With table.ListRows(i)
                .Range(table.ListColumns("TBA + Polar + Statbotics").Index).Value = (.Range(table.ListColumns("Statbotics").Index).Value + .Range(table.ListColumns("Polar").Index).Value + .Range(table.ListColumns("TBA").Index).Value) / 3
                .Range(table.ListColumns("Polar + Statbotics").Index).Value = (.Range(table.ListColumns("Statbotics").Index).Value + .Range(table.ListColumns("Polar").Index).Value) / 2
                .Range(table.ListColumns("TBA + Polar").Index).Value = (.Range(table.ListColumns("TBA").Index).Value + .Range(table.ListColumns("Polar").Index).Value) / 2
                .Range(table.ListColumns("TBA + Statbotics").Index).Value = (.Range(table.ListColumns("Statbotics").Index).Value + .Range(table.ListColumns("TBA").Index).Value) / 2
            End With
        Next team
End Sub
Function GetPolarForecastData(eventKey As String, team As Integer) As Object
    ' Define variables
    Dim requestUrl As String
    Dim http As New MSXML2.XMLHTTP
    Dim responseText As String
    Dim json As Object
    
    ' Construct request URL
    requestUrl = "https://polarforecast.azurewebsites.net/2023/" & eventKey & "/frc" & team & "/stats"
    
    ' Make HTTP request
    http.Open "GET", requestUrl, False
    http.send
    
    ' Get response text
    responseText = http.responseText
    'MsgBox responseText
    
    ' Parse response text as JSON object
    Set json = JsonConverter.ParseJson(responseText)
    
    ' Return JSON object
    Set GetPolarForecastData = json
End Function
Function GetTeams(eventKey As String) As Object
    ' Define variables
    Dim requestUrl As String
    Dim http As New MSXML2.XMLHTTP
    Dim responseText As String
    Dim json As Object
    
    ' Construct request URL
    requestUrl = "https://www.thebluealliance.com/api/v3/event/" & eventKey & "/teams"
    
    ' Make HTTP request
    http.Open "GET", requestUrl, False
    http.setRequestHeader "X-TBA-Auth-Key", "mRpZzqVWf2fc5UNyrzb7UChKwh9edXKlJHEdeE5L5J7jX6BCGmBHIhnCmkZDHnMC"
    http.send
    
    ' Get response text
    responseText = http.responseText
    'MsgBox responseText
    
    ' Parse response text as JSON object
    Set json = JsonConverter.ParseJson(responseText)
    
    ' Return JSON object
    Set GetTeams = json
End Function
Function GetOPRs(eventKey As String) As Object
    ' Define variables
    Dim requestUrl As String
    Dim http As New MSXML2.XMLHTTP
    Dim responseText As String
    Dim json As Object
    
    ' Construct request URL
    requestUrl = "https://www.thebluealliance.com/api/v3/event/" & eventKey & "/oprs"
    
    ' Make HTTP request
    http.Open "GET", requestUrl, False
    http.setRequestHeader "X-TBA-Auth-Key", "mRpZzqVWf2fc5UNyrzb7UChKwh9edXKlJHEdeE5L5J7jX6BCGmBHIhnCmkZDHnMC"
    http.send
    
    ' Get response text
    responseText = http.responseText
    'MsgBox responseText
    
    ' Parse response text as JSON object
    Set json = JsonConverter.ParseJson(responseText)
    
    ' Return JSON object
    Set GetOPRs = json
End Function
Function GetStatboticsData(eventKey As String, team As Integer) As Object
    ' Define variables
    Dim requestUrl As String
    Dim http As New MSXML2.XMLHTTP
    Dim responseText As String
    Dim json As Object
    
    ' Construct request URL
    requestUrl = "https://api.statbotics.io/v2/team_event/" & team & "/" & eventKey
    
    ' Make HTTP request
    http.Open "GET", requestUrl, False
    http.send
    
    ' Get response text
    responseText = http.responseText
    'MsgBox responseText
    
    ' Parse response text as JSON object
    Set json = JsonConverter.ParseJson(responseText)
    
    ' Return JSON object
    Set GetStatboticsData = json
End Function
Sub highlightEntries()
    Dim tableexists As Boolean, max As Integer
    Dim tablename As String, table As ListObject
    Dim row As ListRow
    tablename = "ScoutingData"
    tableexists = False
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim x As Integer, y As Integer
    'Loop through each sheet and table in the workbook
    For Each sht In ThisWorkbook.Worksheets
        For Each tbl In sht.ListObjects
            If tbl.Name = tablename Then
                tableexists = True
                Set table = tbl
                Set ws = sht
            End If
        Next tbl
    Next sht
    If tableexists Then
        'Set table = ws.ListObjects(tableName)
    Else
        MsgBox ("No Table Found")
        Exit Sub
    End If
    For Each row In table.ListRows
        y = row.Range(table.ListColumns("matchNumber").Index).Value Mod 5
                Select Case (y)
                    Case 0:
                        row.Range.Borders.Color = RGB(255, 255, 102)
                        row.Range.Interior.Color = RGB(255, 255, 102)
                    Case 1:
                        row.Range.Borders.Color = RGB(255, 178, 102)
                        row.Range.Interior.Color = RGB(255, 178, 102)
                    Case 2:
                        row.Range.Borders.Color = RGB(102, 178, 255)
                        row.Range.Interior.Color = RGB(102, 178, 255)
                    Case 3:
                        row.Range.Borders.Color = RGB(102, 255, 102)
                        row.Range.Interior.Color = RGB(102, 255, 102)
                    Case 4:
                        row.Range.Borders.Color = RGB(255, 153, 255)
                        row.Range.Interior.Color = RGB(255, 153, 255)
                End Select
    Next row
End Sub
Sub duplicateStations()
    Dim tableexists As Boolean, max As Integer
    Dim tablename As String, table As ListObject
    Dim rows() As ListRow
    tablename = "ScoutingData"
    tableexists = False
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim x As Integer, y As Integer
    'Loop through each sheet and table in the workbook
    For Each sht In ThisWorkbook.Worksheets
        For Each tbl In sht.ListObjects
            If tbl.Name = tablename Then
                tableexists = True
                Set table = tbl
                Set ws = sht
            End If
        Next tbl
    Next sht
    If tableexists Then
        'Set table = ws.ListObjects(tableName)
    Else
        MsgBox ("No Table Found")
        Exit Sub
    End If
    Dim row As ListRow, checkRow As ListRow
    For Each row In table.ListRows
        For Each checkRow In table.ListRows
            If Not checkRow.Range.Address = row.Range.Address Then
                If checkRow.Range(table.ListColumns("matchNumber").Index).Value = row.Range(table.ListColumns("matchNumber").Index).Value Then
                    If checkRow.Range(table.ListColumns("robot").Index).Value = row.Range(table.ListColumns("robot").Index).Value Then
                        row.Range(table.ListColumns("robot").Index).Interior.Color = RGB(255, 49, 49)
                        row.Range(table.ListColumns("robot").Index).Borders.Color = RGB(255, 49, 49)
                    End If
                    If checkRow.Range(table.ListColumns("teamNumber").Index).Value = row.Range(table.ListColumns("teamNumber").Index).Value Then
                        row.Range(table.ListColumns("teamNumber").Index).Interior.Color = RGB(255, 49, 49)
                        row.Range(table.ListColumns("teamNumber").Index).Borders.Color = RGB(255, 49, 49)
                    End If
                End If
            End If
        Next checkRow
    Next row
End Sub
Sub checkNumEntries()
    Dim tableexists As Boolean, max As Integer, z As Range, a As Range
    Dim tablename As String, table As ListObject
    Dim rows() As ListRow
    tablename = "ScoutingData"
    tableexists = False
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim x As Integer, y As Integer
    'Loop through each sheet and table in the workbook
    For Each sht In ThisWorkbook.Worksheets
        For Each tbl In sht.ListObjects
            If tbl.Name = tablename Then
                tableexists = True
                Set table = tbl
                Set ws = sht
            End If
        Next tbl
    Next sht
    If tableexists Then
        'Set table = ws.ListObjects(tableName)
    Else
        MsgBox ("No Table Found")
        Exit Sub
    End If
    Dim row As ListRow, checkRow As ListRow
    max = Application.WorksheetFunction.max(table.ListColumns("matchNumber").Range)
    For Each row In table.ListRows
        For Each checkRow In table.ListRows
            If checkRow.Range(table.ListColumns("matchNumber").Index).Value = row.Range(table.ListColumns("matchNumber").Index).Value Then
                x = x + 1
            End If
        Next checkRow
        If Not x = 6 Then
            row.Range(table.ListColumns("matchNumber").Index).Interior.Color = RGB(255, 49, 49)
            row.Range(table.ListColumns("matchNumber").Index).Borders.Color = RGB(255, 49, 49)
        End If
        x = 0
    Next row
End Sub
Function writeTeams()
    Dim row As Integer, rows As Integer, team, checkRow As Integer, switches As Integer, hold As Variant, temp As Variant
    Worksheets("Numerical").Range("B2:B" & (numRows("Numerical") - 1)).copy Worksheets("Average").Range("A2")
    rows = numRows("Average") + 1
    For row = 2 To rows
        team = Worksheets("Average").Range("A" & row).Value
        For checkRow = row + 1 To rows
            If Worksheets("Average").Range("A" & checkRow).Value = team Then
                Worksheets("Average").Range("A" & checkRow).Value = Null
            End If
        Next checkRow
    Next row
    switches = 1
    Do While Not switches = 0
        switches = 0
        For row = 2 To rows
            If Worksheets("Average").Range("A" & row).Value < Worksheets("Average").Range("A" & (row + 1)).Value Then
                hold = Worksheets("Average").Range("A" & row).Value
                temp = Worksheets("Average").Range("A" & (row + 1)).Value
                Worksheets("Average").Range("A" & row).Value = temp
                Worksheets("Average").Range("A" & (row + 1)).Value = hold
                switches = switches + 1
            End If
        Next row
    Loop
End Function
Function averageColumn(column As Integer) As Integer
    Dim row As Integer, val As Double, div As Integer, team, x
    For row = 2 To numRows("Average") - 1
        val = 0
        div = 0
        team = Worksheets("Average").Range("A" & row).Value
        For x = 2 To numRows("Numerical")
            If Worksheets("Numerical").Range("B" & x).Value = team Then
                If Not Worksheets("Numerical").Range(columnLetter(column) & x).Value < 0 Then
                        val = val + Worksheets("Numerical").Range(columnLetter(column) & x).Value
                        div = div + 1
                End If
            End If
        Next x
        If div = 0 Then
            Worksheets("Average").Range(columnLetter(column - 1) & row).Value = 0
        Else
            Worksheets("Average").Range(columnLetter(column - 1) & row).Value = val / div
        End If
    Next row
End Function
Function hasCard(column As Integer) As Integer
    Dim row As Integer, val As Boolean, team, x
    For row = 2 To numRows("Average") - 1
        val = False
        team = Worksheets("Average").Range("A" & row).Value
        For x = 2 To numRows("Numerical")
            If Worksheets("Numerical").Range("B" & x).Value = team Then
                If Worksheets("Numerical").Range(columnLetter(column) & x).Value = 1 Then
                    val = True
                End If
            End If
        Next x
        Worksheets("Average").Range(columnLetter(column - 1) & row).Value = val
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
Function contribution(points As Integer, alliancePoints As Integer, sendto As Integer, row As Integer) As Integer
    Dim val As Double, sheet As Worksheet, p As Double, w As Double
    Set sheet = Worksheets("Numerical")
    p = sheet.Range(columnLetter(points) & row).Value
    w = sheet.Range(columnLetter(alliancePoints) & row).Value
    val = p / w
    sheet.Range(columnLetter(sendto) & row).Value = val
    contribution = sendto + 1
End Function
Function copy(getFrom As Integer, sendto As Integer, row As Integer) As Integer
    Dim val As Variant
    val = Worksheets("Input").Range(columnLetter(getFrom) & row).Value
    Worksheets("Numerical").Range(columnLetter(sendto) & row).Value = val
    copy = sendto + 1
End Function
Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
Function columnLetter(columnNumber As Integer) As String
    columnLetter = Split(Cells(1, columnNumber).Address, "$")(1)
End Function
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
    checkErrors
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
    Dim tablename As String
    tablename = "ScoutingData"
    ' Set up map
    ' Fields for every year
    mapper.Add "s", "scouter"
    mapper.Add "e", "eventCode"
    mapper.Add "l", "matchLevel"
    mapper.Add "m", "matchNumber"
    mapper.Add "r", "robot"
    mapper.Add "t", "teamNumber"
    mapper.Add "sp", "StartPosition"
    mapper.Add "ahc", "autoHighCones"
    mapper.Add "ahcu", "autoHighCubes"
    mapper.Add "amc", "autoMidCones"
    mapper.Add "amcu", "autoMidCubes"
    mapper.Add "alc", "autoLowPieces"
    mapper.Add "ad", "autoDocking"
    mapper.Add "as", "autoStrategy"
    mapper.Add "pm", "piecesMissed"
    mapper.Add "ec", "exitedCommunity"
    mapper.Add "hc", "teleopHighCones"
    mapper.Add "hcu", "teleopHighCubes"
    mapper.Add "mc", "teleopMidCones"
    mapper.Add "mcu", "teleopMidCubes"
    mapper.Add "lc", "teleopLowPieces"
    mapper.Add "ts", "teleopStrategy"
    mapper.Add "dt", "dockingTimer"
    mapper.Add "fs", "finalStatus"
    mapper.Add "stg", "struggled"
    mapper.Add "dn", "totalDockedBots"
    mapper.Add "es", "endgameStrategy"
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
    mapper.Add "gang", "allianceScore"
    mapper.Add "opp", "oppositionScore"
    'Additional custom mapping
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
                If tbl.Name = tablename Then
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
            ws.ListObjects.Add(xlSrcRange, Range("A1:AO1"), , xlYes).Name = tablename
            i = 0
            Set table = ws.ListObjects(tablename)
            For Each Key In data.Keys
                table.Range(i + 1) = Key
                i = i + 1
            Next
        End If
        Dim newrow As ListRow
        Set newrow = table.ListRows.Add
        For Each str In data.Keys
            If str = "driverSkill" Or str = "defenseRating" Then
                If data(str) = 0 Then
                    newrow.Range(table.ListColumns(str).Index) = data(str) - 1
                Else
                    newrow.Range(table.ListColumns(str).Index) = data(str)
                End If
            Else
                newrow.Range(table.ListColumns(str).Index) = data(str)
            End If
        Next
        Dim x As Integer
        x = newrow.Range(table.ListColumns("matchNumber").Index).Value Mod 5
        Select Case (x)
            Case 0:
                newrow.Range.Interior.Color = RGB(255, 255, 102)
            Case 1:
                newrow.Range.Interior.Color = RGB(255, 178, 102)
            Case 2:
                newrow.Range.Interior.Color = RGB(102, 178, 255)
            Case 3:
                newrow.Range.Interior.Color = RGB(102, 255, 102)
            Case 4:
                newrow.Range.Interior.Color = RGB(255, 153, 255)
        End Select
    End If
End Sub
Sub SecondPick()
    Dim sheet As String, row As Integer, sendto As Integer, val As Double, column As Integer, weightsFrom As Integer, x As Integer
    sheet = "average"
    sendto = 31
    weightsFrom = CInt(InputBox("From which row are you setting the weights?"))
    For x = 1 To CInt(InputBox("How many weighted scores?"))
    For row = 2 To numRows(sheet) - 1
        val = 0
        For column = 2 To 30
            val = val + Worksheets(sheet).Range(columnLetter(column) & weightsFrom).Value * Worksheets("Average").Range(columnLetter(column) & row).Value
        Next column
        Worksheets(sheet).Range(columnLetter(sendto) & row).Value = val
    Next row
    sendto = sendto + 1
    weightsFrom = weightsFrom + 1
    Next x
End Sub