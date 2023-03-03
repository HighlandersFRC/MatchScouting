Sub AggregateData()
    teamNumber
    allGamePieces
    teleopDocking
    autoDocking
    startPosition
    cycles
    exitedCommunity
    stillHasAutoPiece
    totalDocked
    died
    skill
    points
    fouls
    writeTeams
    averageColumns
End Sub
Sub totalDocked()
    Dim x As Integer, bots
    For x = 2 To numRows("ScoutingPASS_Excel_Example") - 1
        bots = Worksheets("ScoutingPASS_Excel_Example").Range("S" & x).Value
        Worksheets("Aggregate_Data").Range("T" & x).Value = bots
    Next x
End Sub
'Sub syncSkill()
 '   Dim row As Integer, teamRow As Integer, hasTeam As Boolean, team As Integer, divisor As Integer, val As Double, val1 As Double
  '  For row = 2 To numRows("Skill")
   '     hasTeam = False
    '    divisor = 0
     '   val = 0
      '  val1 = 0
       ' For teamRow = 2 To numRows("ByTeamAverageData")
        '    If Worksheets("ByTeamAverageData").Range("A" & teamRow).Value = Worksheets("Skill").Range("B" & row).Value Then
         '       hasTeam = True
          '      teamRow = teamRow - 1
           '     Exit For
            'End If
'        Next teamRow
 '       For team = row + 1 To numRows("Skill")
  '          If Worksheets("Skill").Range("B" & teamRow).Value = Worksheets("Skill").Range("B" & team).Value Then
   '             divisor = divisor + 1
    '            val = val + Worksheets("Skill").Range("C" & team).Value
     '           val1 = val1 + Worksheets("Skill").Range("D" & team).Value
      '      End If
       ' Next team
        'If hasTeam Then
'        If Not divisor = 0 Then
 '           Worksheets("ByTeamAverageData").Range("R" & teamRow).Value = val / divisor
  '          Worksheets("ByTeamAverageData").Range("S" & teamRow).Value = val1 / divisor
   '     End If
    '    End If
'    Next row
'End Sub
Sub syncPit()
    Dim rows As Integer, teamRow, hasTeam As Boolean, team, rng
    For rows = 2 To numRows("PitScouting")
        hasTeam = False
        team = Worksheets("PitScouting").Range("B" & rows).Value
        For teamRow = 2 To numRows("ByTeamAverageData")
            If Worksheets("ByTeamAverageData").Range("A" & teamRow).Value = team Then
                hasTeam = True
                
                Exit For
            End If
        Next teamRow
    If hasTeam Then
        rng = Sheets("PitScouting").Range("A" & rows & ":V" & rows)
        Sheets("ByTeamAverageData").Range("W" & teamRow & ":AP" & teamRow) = rng
    End If
    Next rows
End Sub
Sub fouls()
    Dim row As Integer, fouls
    For row = 2 to numRows("ScoutingPASS_Excel_Example") -1
        fouls = Worksheets("Scouting_PASS_Excel_Example").Range("Y" & row).Value
        Worksheets("Aggregate_Data").Range("V" & row).Value = fouls
    Next row
End Sub
Sub sortByColumn(Column As String, Worksheet As String)
    Dim row As Integer, hold1 As Variant, hold2 As Variant, switches As Integer
    switches = 1
    Do While Not switches = 0
        switches = 0
        For row = 2 To numRows(Worksheet) - 2
            If Worksheets(Worksheet).Range(Column & row).Value < Worksheets(Worksheet).Range(Column & (row + 1)).Value Then
                hold1 = Worksheets(Worksheet).Range("A" & row & ":Z" & row)
                hold2 = Worksheets(Worksheet).Range("A" & (row + 1) & ":AZ" & (row + 1))
                Worksheets(Worksheet).Range("A" & row & ":Z" & row) = hold2
                Worksheets(Worksheet).Range("A" & (row + 1) & ":Z" & (row + 1)) = hold1
                switches = switches + 1
            End If
        Next row
    Loop
End Sub
Sub sortByColumnInverse(Column As String, Worksheet As String)
    Dim row As Integer, hold1 As Variant, hold2 As Variant, switches As Integer
    switches = 1
    Do While Not switches = 0
        switches = 0
        For row = 2 To numRows(Worksheet) - 2
            If Worksheets(Worksheet).Range(Column & row).Value > Worksheets(Worksheet).Range(Column & (row + 1)).Value Then
                hold1 = Worksheets(Worksheet).Range("A" & row & ":Z" & row)
                hold2 = Worksheets(Worksheet).Range("A" & (row + 1) & ":AZ" & (row + 1))
                Worksheets(Worksheet).Range("A" & row & ":Z" & row) = hold2
                Worksheets(Worksheet).Range("A" & (row + 1) & ":Z" & (row + 1)) = hold1
                switches = switches + 1
            End If
        Next row
    Loop
End Sub
Sub skill()
    Dim skill, defense, row
    For row = 2 To numRows("ScoutingPASS_Excel_Example") - 1
        skill = Worksheets("ScoutingPASS_Excel_Example").Range("U" & row).Value
        defense = Worksheets("ScoutingPASS_Excel_Example").Range("V" & row).Value
        Select Case skill
            Case "x"
                skill = -1
            Case more
                skill = 3
            Case "a"
                skill = 2
            Case "l"
                skill = 1
            Case Else
                skill = -1
        End Select
        Select Case defense
            Case "x"
                defense = 0
            Case "g"
                defense = 3
            Case "a"
                defense = 2
            Case "b"
                defense = 1
            Case "e"
                defense = 4
            Case Else
                defense = 0
        End Select
        Worksheets("Aggregate_Data").Range("R" & row).Value = skill
        Worksheets("Aggregate_Data").Range("S" & row).Value = defense
    Next row
End Sub
Sub writeTeams()
    Dim row As Integer, rows As Integer, team, checkRow As Integer, switches As Integer, hold As Variant, temp As Variant
    For row = 2 To numRows("Aggregate_Data")
        Worksheets("ByTeamAverageData").Range("A" & row) = Worksheets("Aggregate_Data").Range("A" & row)
    Next row
    rows = numRows("ByTeamAverageData")
    For row = 2 To rows
        team = Worksheets("ByTeamAverageData").Range("A" & row).Value
        For checkRow = row + 1 To rows
            If Worksheets("ByTeamAverageData").Range("A" & checkRow).Value = team Then
                Worksheets("ByTeamAverageData").Range("A" & checkRow).Value = Null
            End If
        Next checkRow
    Next row
    switches = 1
    Do While Not switches = 0
        switches = 0
        For row = 2 To rows
            If Worksheets("ByTeamAverageData").Range("A" & row).Value < Worksheets("ByTeamAverageData").Range("A" & (row + 1)).Value Then
                hold = Worksheets("ByTeamAverageData").Range("A" & row).Value
                temp = Worksheets("ByTeamAverageData").Range("A" & (row + 1)).Value
                Worksheets("ByTeamAverageData").Range("A" & row).Value = temp
                Worksheets("ByTeamAverageData").Range("A" & (row + 1)).Value = hold
                switches = switches + 1
            End If
        Next row
    Loop
End Sub
Sub points()
    Dim row As Integer, weights() As Double, points As Double, switches As Integer, hold As Variant, temp As Variant
    ReDim weights(12)
    switches = 1
    weights(0) = 6
    weights(1) = 6
    weights(2) = 4
    weights(3) = 4
    weights(4) = 3
    weights(5) = 8
    weights(6) = 5
    weights(7) = 5
    weights(8) = 3
    weights(9) = 3
    weights(10) = 2
    weights(11) = 6
    For row = 2 To numRows("Aggregate_Data") - 1
        points = 0
        points = points + weights(0) * Worksheets("Aggregate_Data").Range("B" & row).Value
        points = points + weights(1) * Worksheets("Aggregate_Data").Range("C" & row).Value
        points = points + weights(2) * Worksheets("Aggregate_Data").Range("D" & row).Value
        points = points + weights(3) * Worksheets("Aggregate_Data").Range("E" & row).Value
        points = points + weights(4) * Worksheets("Aggregate_Data").Range("F" & row).Value
        points = points + weights(5) * Worksheets("Aggregate_Data").Range("G" & row).Value
        points = points + weights(6) * Worksheets("Aggregate_Data").Range("J" & row).Value
        points = points + weights(7) * Worksheets("Aggregate_Data").Range("K" & row).Value
        points = points + weights(8) * Worksheets("Aggregate_Data").Range("L" & row).Value
        points = points + weights(9) * Worksheets("Aggregate_Data").Range("M" & row).Value
        points = points + weights(10) * Worksheets("Aggregate_Data").Range("N" & row).Value
        points = points + weights(11) * Worksheets("Aggregate_Data").Range("P" & row).Value
        Worksheets("Aggregate_Data").Range("U" & row).Value = points
    Next row
End Sub
Sub averageColumns()
    averageColumn ("B")
    averageColumn ("C")
    averageColumn ("D")
    averageColumn ("E")
    averageColumn ("F")
    averageColumn ("G")
    averageColumn ("H")
    averageColumn ("I")
    averageColumn ("J")
    averageColumn ("K")
    averageColumn ("L")
    averageColumn ("M")
    averageColumn ("N")
    averageColumn ("O")
    averageColumn ("P")
    averageColumn ("Q")
    averageColumn ("R")
    averageColumn ("S")
    averageColumn ("T")
    averageColumn ("U")
    averageColumn ("V")
End Sub
Sub averageColumn(Column As String)
    Dim row As Integer, team, teamRow As Integer, y As Integer, z As Integer, Value
    For row = 2 To numRows("Aggregate_Data") - 1
        team = Worksheets("Aggregate_Data").Range("A" & row).Value
        For y = 2 To numRows("ByTeamAverageData") - 1
            If team = Worksheets("ByTeamAverageData").Range("A" & y).Value Then
                teamRow = y
                Exit For
            End If
        Next y
        For z = 2 To numRows("Aggregate_Data")
            If Not Column = "Q" Then
                If Column = "H" Then
                    If Not Worksheets("Aggregate_Data").Range(Column & z).Value < 0 Then
                        If Worksheets("Aggregate_Data").Range("A" & z).Value = Worksheets("Aggregate_Data").Range("A" & row).Value Then
                            divisor = divisor + Worksheets("Aggregate_Data").Range("I" & z).Value
                            Value = Value + Worksheets("Aggregate_Data").Range("H" & z).Value * Worksheets("Aggregate_Data").Range("I" & z).Value
                        End If
                    End If
                Else
                    If Not Worksheets("Aggregate_Data").Range(Column & z).Value < 0 Then
                        If Worksheets("Aggregate_Data").Range("A" & z).Value = Worksheets("Aggregate_Data").Range("A" & row).Value Then
                            divisor = divisor + 1
                            Value = Value + Worksheets("Aggregate_Data").Range(Column & z).Value
                        End If
                    End If
                End If
            Else
                If Not Worksheets("Aggregate_Data").Range(Column & z).Value < 0 Then
                    If Worksheets("Aggregate_Data").Range("A" & z).Value = Worksheets("Aggregate_Data").Range("A" & row).Value Then
                        divisor = divisor + 1
                        Value = Value + Worksheets("Aggregate_Data").Range(Column & z).Value
                    End If
                End If
            End If
        Next z
        If Not divisor = 0 Then
            If Not Worksheets("ByTeamAverageData").Range("A" & teamRow).Value = "" Then
                Worksheets("ByTeamAverageData").Range(Column & teamRow).Value = Value / divisor
            End If
        Else
            If Not Worksheets("ByTeamAverageData").Range("A" & teamRow).Value = "" Then
                Worksheets("ByTeamAverageData").Range(Column & teamRow).Value = 0
            End If
        End If
        divisor = 0
        Value = 0
    Next row
End Sub
Sub teamNumber()
    Dim x As Integer
    For x = 1 To numRows("ScoutingPASS_Excel_Example") + 1
        Worksheets("Aggregate_Data").Range("A" & x).Value = Worksheets("ScoutingPASS_Excel_Example").Range("E" & x).Value
        Worksheets("Autos").Range("A" & x).Value = Worksheets("ScoutingPASS_Excel_Example").Range("E" & x).Value
    Next x
End Sub
Sub exitedCommunity()
    Dim bool As Boolean
    bool = tf("I", "D", "Autos")
End Sub
Sub died()
    Dim bool As Boolean
    bool = tf("W", "Q", "Aggregate_Data")
End Sub
Sub stillHasAutoPiece()
    Dim bool As Boolean
    bool = tf("L", "E", "Autos")
End Sub
Function tf(getFromColumn As String, mapToColumn As String, mapToWorkSheet As String) As Boolean
    Dim x As Integer, truefalse As Boolean
    For x = 2 To numRows("ScoutingPASS_Excel_Example") - 1
        Worksheets(mapToWorkSheet).Range(mapToColumn & x).Value = Worksheets("ScoutingPASS_Excel_Example").Range(getFromColumn & x).Value
    Next x
End Function
Sub WriteEventData()
    Dim url As String
    Dim http As New MSXML2.XMLHTTP60
    Dim response As String
    Dim jsonMatches As Object
    Dim match As Integer
    
    Dim currentRow As Long
    currentRow = 1 ' Change this to the row where you want to start adding data
    
    Dim currentColumn As Long
    currentColumn = 1 ' Change this to the column where you want to start adding data
    
    Dim currentMatch As String
    Set jsonMatches = JsonConverter.ParseJson(GetTeamData("2023week0"))
    For match = 1 To 1000
        MsgBox (match)
        If IsNull(jsonMatches(match)) Then
            Exit For
        End If
        currentRow = currentRow + 1
        Worksheets("BlueAlliance").Range("A" & currentRow).Value = jsonMatches(match)("key")
        For currentColumn = 2 To 4
            Worksheets("BlueAlliance").Range(ColumnLetter(currentColumn) & currentRow).Value = jsonMatches(match)("alliances")("blue")("team_keys")(currentColumn - 1)
        Next currentColumn
        For currentColumn = 5 To 7
            Worksheets("BlueAlliance").Range(ColumnLetter(currentColumn) & currentRow).Value = jsonMatches(match)("alliances")("red")("team_keys")(currentColumn - 4)
        Next currentColumn
    Next match
End Sub
Sub WriteOPRs()
    Dim rankings As Object, json As Object, currentColumn As Long, currentRow As Long, team As Integer, eventCode As String
    eventCode = InputBox("event code")
    Set json = JsonConverter.ParseJson(GetEventOPRs(eventCode))
    For currentRow = 2 To numRows("TeamRankings")
        If Not IsNull(json("oprs")(currentRow - 1)) Then
            Exit For
        End If
        currentColumn = 7
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow).Value = json("oprs")("frc" + CStr(Worksheets("TeamRankings").Range("A" & currentRow).Value))
        currentColumn = currentColumn + 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow).Value = json("dprs")("frc" + CStr(Worksheets("TeamRankings").Range("A" & currentRow).Value))
        currentColumn = currentColumn + 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow).Value = json("ccwms")("frc" + CStr(Worksheets("TeamRankings").Range("A" & currentRow).Value))
        currentColumn = currentColumn + 1
    Next currentRow
End Sub
Sub writeTeamEventData()
    Dim rankings As Object, json As Object, currentColumn As Long, currentRow As Long, team As Integer, eventCode As String
    eventCode = InputBox("Event Code")
    Set json = JsonConverter.ParseJson(GetRankings(eventCode))
    currentColumn = 1
    currentRow = 1
    On Error GoTo oprs
    For team = 1 To 2000
        If IsNull(json("rankings")(team)) Then
            Exit For
        End If
        currentRow = currentRow + 1
        currentColumn = 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow) = Replace(json("rankings")(team)("team_key"), "frc", "")
        currentColumn = currentColumn + 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow) = json("rankings")(team)("rank")
        currentColumn = currentColumn + 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow) = json("rankings")(team)("sort_orders")(1)
        currentColumn = currentColumn + 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow) = json("rankings")(team)("sort_orders")(2)
        currentColumn = currentColumn + 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow) = json("rankings")(team)("sort_orders")(3)
        currentColumn = currentColumn + 1
        Worksheets("TeamRankings").Range(ColumnLetter(currentColumn) & currentRow) = json("rankings")(team)("sort_orders")(4)
    Next team
oprs:
End Sub
Function ColumnLetter(columnNumber As Long) As String
    ' Convert a column number to its corresponding letter
    ColumnLetter = Split(Cells(1, columnNumber).Address, "$")(1)
End Function
Function GetRankings(eventCode As String) As String
    Dim url As String
    Dim http As New MSXML2.XMLHTTP60
    Dim response As String
    Dim jsonMatches As Object
    url = "https://www.thebluealliance.com/api/v3/event/" & eventCode & "/rankings"
    http.Open "GET", url, False
    http.setRequestHeader "X-TBA-Auth-Key", "puxQygNBxY7TOUjDhyddDsXFmFZORieMOGsWh0cG66tkGXrfd18DkKQwgwA5wyFz"
    http.send
    GetRankings = http.responseText
    MsgBox (GetRankings)
End Function
Function GetEventOPRs(eventCode As String) As String
    Dim url As String
    Dim http As New MSXML2.XMLHTTP60
    Dim response As String
    Dim jsonMatches As Object
    url = "https://www.thebluealliance.com/api/v3/event/" & eventCode & "/oprs"
    http.Open "GET", url, False
    http.setRequestHeader "X-TBA-Auth-Key", "puxQygNBxY7TOUjDhyddDsXFmFZORieMOGsWh0cG66tkGXrfd18DkKQwgwA5wyFz"
    http.send
    GetEventOPRs = http.responseText
    MsgBox (GetEventOPRs)
End Function
Function GetTeamData(teamNumber As String) As String
    Dim apiKey As String
    Dim url As String
    Dim http As New MSXML2.XMLHTTP60
    
    apiKey = "puxQygNBxY7TOUjDhyddDsXFmFZORieMOGsWh0cG66tkGXrfd18DkKQwgwA5wyFz"
    url = "https://www.thebluealliance.com/api/v3/team/" & teamNumber
    
    http.Open "GET", url, False
    http.setRequestHeader "X-TBA-Auth-Key", apiKey
    http.send
    
    GetTeamData = http.responseText
End Function
Sub startPosition()
    Dim x As Integer, alliance As String, startPos
    For x = 2 To numRows("ScoutingPASS_Excel_Example") - 1
        alliance = Worksheets("ScoutingPASS_Excel_Example").Range("E" & x).Value
        If Not InStr(1, alliance, "b") = 0 Then
            alliance = "b"
        Else
            alliance = "r"
        End If
        startPos = Worksheets("ScoutingPASS_Excel_Example").Range("H" & x)
        Worksheets("Autos").Range("B" & x).Value = alliance
        Worksheets("Autos").Range("C" & x).Value = startPos
    Next x
End Sub
Sub autoDocking()
    Dim x As Integer
    For x = 2 To numRows("ScoutingPASS_Excel_Example") - 1
        If Not Worksheets("ScoutingPASS_Excel_Example").Range("J" & x).Value = "x" Then
            If Worksheets("ScoutingPASS_Excel_Example").Range("J" & x).Value = "e" Then
                Worksheets("Autos").Range("K" & x).Value = 1.5
                Worksheets("Aggregate_Data").Range("G" & x).Value = 1.5
            Else
                Worksheets("Autos").Range("K" & x).Value = 1
                Worksheets("Aggregate_Data").Range("G" & x).Value = 1
            End If
        Else
            Worksheets("Autos").Range("K" & x).Value = 0
            Worksheets("Aggregate_Data").Range("G" & x).Value = 0
        End If
    Next x
End Sub
Sub cycles()
    Dim cyclesString() As String
    Dim numCycles As Integer
    Dim average As Double
    Dim x As Integer
    Dim y As Integer
    For x = 2 To numRows("ScoutingPASS_Excel_Example") - 1
        cyclesString() = Split(Worksheets("ScoutingPASS_Excel_Example").Range("M" & x), ",")
        numCycles = UBound(cyclesString) - LBound(cyclesString) + 1
        If Not numCycles = 0 Then
            For y = 0 To numCycles - 1
                If cyclesString(y) = "[]" Then
                    numCycles = 0
                    Exit For
                End If
                cyclesString(y) = Replace(cyclesString(y), ",", "")
                cyclesString(y) = Replace(cyclesString(y), """", "")
                cyclesString(y) = Replace(cyclesString(y), "[", "")
                cyclesString(y) = Replace(cyclesString(y), "]", "")
                average = average + CDbl(cyclesString(y))
            Next
            If numCycles = 0 Then
                average = 0
            Else
                average = average / (numCycles * 2) / numCycles
            End If
        Else
            average = 0
        End If
        Worksheets("Aggregate_Data").Range("H" & x).Value = average
        Worksheets("Aggregate_Data").Range("I" & x).Value = numCycles
    Next x
End Sub
Sub teleopDocking()
    Dim x As Integer, status
    For x = 2 To numRows("ScoutingPASS_Excel_Example") - 1
        If Not Worksheets("ScoutingPASS_Excel_Example").Range("R" & x).Value = "x" Then
            Worksheets("Aggregate_Data").Range("O" & x).Value = Worksheets("ScoutingPASS_Excel_Example").Range("Q" & x).Value
            If Worksheets("ScoutingPASS_Excel_Example").Range("R" & x).Value = "e" Then
                Worksheets("Aggregate_Data").Range("P" & x).Value = 5 / 3
            Else
                If Worksheets("ScoutingPASS_Excel_Example").Range("R" & x).Value = "p" Then
                    Worksheets("Aggregate_Data").Range("P" & x).Value = 1 / 3
                Else
                    Worksheets("Aggregate_Data").Range("P" & x).Value = 1
                End If
            End If
        Else
            Worksheets("Aggregate_Data").Range("O" & x).Value = 0
            Worksheets("Aggregate_Data").Range("P" & x).Value = 0
        End If
    Next x
End Sub
Sub allGamePieces()
    Dim x As Integer
    For x = 1 To (numRows("ScoutingPASS_Excel_Example") - 2)
        AutoGamePieces (x + 1)
        TeleopGamePieces (x + 1)
        Application.ScreenUpdating = True
    Next x
End Sub
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
Function AutoGamePieces(row As String)
    Dim x As Integer, y As Integer
    Dim strTest As String
    Dim numHighCubes As Integer
    Dim numHighCones As Integer
    Dim numMidCubes As Integer
    Dim numMidCones As Integer
    Dim numLowPieces As Integer
    Dim strString() As String
    Dim strDouble() As Double
    Dim numPieces As Integer
    numHighCubes = 0
    numHighCones = 0
    numMidCubes = 0
    numMidCones = 0
    strTest = Worksheets("ScoutingPASS_Excel_Example").Range("G" & row).Value
    strString() = Split(strTest, ",")
    numPieces = UBound(strString) - LBound(strString) + 1
    ReDim strDouble(numPieces)
    If Not numPieces = 0 Then
        For x = 0 To numPieces - 1
            If strString(x) = "[]" Then
                numPieces = 0
                Exit For
            End If
            strString(x) = Replace(strString(x), ",", "")
            strString(x) = Replace(strString(x), """", "")
            strString(x) = Replace(strString(x), "[", "")
            strString(x) = Replace(strString(x), "]", "")
            strDouble(x) = CDbl(strString(x))
        Next
    End If
    If Not numPieces = 0 Then
    For x = 0 To numPieces - 1
        For y = 1 To 27
            If strDouble(x) = y Then
                Exit For
            End If
        Next y
        If y < 19 Then
            If (y + 1) Mod 3 = 0 Then
                If y > 9 Then
                    numMidCubes = numMidCubes + 1
                Else
                    numHighCubes = numHighCubes + 1
                End If
            Else
            If y > 9 Then
                numMidCones = numMidCones + 1
            Else
                numHighCones = numHighCones + 1
            End If
            End If
        Else
            numLowPieces = numLowPieces + 1
        End If
    Next x
    End If
    Worksheets("Autos").Range("F" + row).Value = numHighCubes
    Worksheets("Autos").Range("G" + row).Value = numHighCones
    Worksheets("Autos").Range("H" + row).Value = numMidCubes
    Worksheets("Autos").Range("I" + row).Value = numMidCones
    Worksheets("Autos").Range("J" + row).Value = numLowPieces
    Worksheets("Aggregate_Data").Range("B" + row).Value = numHighCubes
    Worksheets("Aggregate_Data").Range("C" + row).Value = numHighCones
    Worksheets("Aggregate_Data").Range("D" + row).Value = numMidCubes
    Worksheets("Aggregate_Data").Range("E" + row).Value = numMidCones
    Worksheets("Aggregate_Data").Range("F" + row).Value = numLowPieces
End Function
Function TeleopGamePieces(row As String)
   Dim x As Integer, y As Integer
    Dim strTest As String
    Dim numHighCubes As Integer
    Dim numHighCones As Integer
    Dim numMidCubes As Integer
    Dim numMidCones As Integer
    Dim numLowPieces As Integer
    Dim strString() As String
    Dim strDouble() As Double
    Dim numPieces As Integer
    numHighCubes = 0
    numHighCones = 0
    numMidCubes = 0
    numMidCones = 0
    strTest = Worksheets("ScoutingPASS_Excel_Example").Range("M" & row).Value
    strString() = Split(strTest, ",")
    numPieces = UBound(strString) - LBound(strString) + 1
    ReDim strDouble(numPieces)
    If Not numPieces = 0 Then
        For x = 0 To numPieces - 1
            If strString(x) = "[]" Then
                numPieces = 0
                Exit For
            End If
            strString(x) = Replace(strString(x), ",", "")
            strString(x) = Replace(strString(x), """", "")
            strString(x) = Replace(strString(x), "[", "")
            strString(x) = Replace(strString(x), "]", "")
            strDouble(x) = CDbl(strString(x))
        Next
    End If
    If Not numPieces = 0 Then
    For x = 0 To numPieces - 1
        For y = 1 To 27
            If strDouble(x) = y Then
                Exit For
            End If
        Next y
        If y < 19 Then
            If (y + 1) Mod 3 = 0 Then
                If y > 9 Then
                    numMidCubes = numMidCubes + 1
                Else
                    numHighCubes = numHighCubes + 1
                End If
            Else
            If y > 9 Then
                numMidCones = numMidCones + 1
            Else
                numHighCones = numHighCones + 1
            End If
            End If
        Else
            numLowPieces = numLowPieces + 1
        End If
    Next x
    End If
    Worksheets("Aggregate_Data").Range("J" + row).Value = numHighCubes
    Worksheets("Aggregate_Data").Range("K" + row).Value = numHighCones
    Worksheets("Aggregate_Data").Range("L" + row).Value = numMidCubes
    Worksheets("Aggregate_Data").Range("M" + row).Value = numMidCones
    Worksheets("Aggregate_Data").Range("N" + row).Value = numLowPieces
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
End Sub

Public Function getInput()
    getInput = InputBox("Scan QR Code", "Match Scouting Input")
End Function

Sub testSaveData()
    saveData ("s=fff;e=1234;l=qm;m=1234;r=r1;t=1234;as=;ae=Y;al=2;ao=2;ai=1;aa=Y;at=N;ax=Y;lp=2;op=1;ip=3;rc=pass;f=0;pc=pass;ss=;c=pass;b=N;ca=x;cb=x;cs=slow;p=N;ds=x;dr=x;pl=x;tr=N;wd=N;if=N;d=N;to=N;be=N;cf=N")
End Sub

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

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
    mapper.Add "ad", "docked"
    mapper.Add "ha", "hasAuto"
    mapper.Add "agpa", "autoAttemptedPieces"
    mapper.Add "gph", "gamePiecesStillWithBot"
    mapper.Add "tct", "Cycles"
    mapper.Add "tsg", "teleopScoring"
    mapper.Add "fo", "fedOthers#Pieces"
    mapper.Add "of", "othersFed#Pieces"
    mapper.Add "dc", "droppedPieces"
    mapper.Add "dt", "dockingTimer"
    mapper.Add "fs", "finalStatus"
    mapper.Add "dn", "totalDockedBots"
    mapper.Add "ds", "driverSkill"
    mapper.Add "dr", "defenseRating"
    mapper.Add "wd", "wasDefended"
    mapper.Add "die", "died/immobilized"
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