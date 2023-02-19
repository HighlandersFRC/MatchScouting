Attribute VB_Name = "QRReader"

Sub Save1QR()
    saveData (getInput())
    ActiveWorkbook.Save
End Sub

Sub processHardCodedData()
    saveData ("s=fudd;e=2022carv;l=qm;m=2;r=r2;t=2451;as=[35];asg=[3,4];acc=1;acs=1;am=1;ad=e;tct=[8.3,7.3,6.7,7.1,5.5,5.8,5.4];tsg=[5,6,7,8,9,1,2];tfc=0;wf=0;wd=0;who=;lnk=1;fpu=b;dt=9.9;fs=e;dn=2;ds=v;ls=5;dr=x;sd=1;sr=5;die=0;tip=0;dc=0;all=1;co=PWNAGE")
    
    ActiveWorkbook.Save
End Sub

Sub processQRCodeInput()
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    ActiveWorkbook.Save
End Sub

Sub Save1PitQR()
    savePitData (getInput())
    ActiveWorkbook.Save
End Sub

Public Function getInput()
    getInput = InputBox("Scan QR Code", "2023 Match Scouting Input")
End Function
'Public Function Scaner()
'    Dim addIn As COMAddIn
'    Dim automationObject As Object
'    Set addIn = Application.COMAddIns("QRReader")
'    Set automationObject = addIn.Object
'    Dim out As String
'    out = automationObject.Scaner
'    Scaner = out
'End Function

Sub test()
    saveData ("s=fudd;e=2022carv;l=qm;m=2;r=r2;t=2451;as=[35];asg=[3,4];acc=1;acs=1;am=1;ad=e;tct=[8.3,7.3,6.7,7.1,5.5,5.8,5.4];tsg=[5,6,7,8,9,1,2];tfc=0;wf=0;wd=0;who=;lnk=1;fpu=b;dt=9.9;fs=e;dn=2;ds=v;ls=5;dr=x;sd=1;sr=5;die=0;tip=0;dc=0;all=1;co=PWNAGE")
End Sub

Sub dbm(inp As String)
    Dim r
    r = MsgBox(inp, vbDefaultButton1 + vbInformation, "Debug", "help.hlp", 1000)
End Sub

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Sub saveData(inp As String)
    Dim fields
    Dim par
    Dim value
    Dim key
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
    mapper.add "asg", "autoScoring"
    mapper.add "ec", "exitedCommunity"
    mapper.add "ad", "docked"
    mapper.add "ha", "hasAuto"
    mapper.add "agpa", "autoAttemptedPieces"
    mapper.add "gph", "gamePiecesStillWithBot"
    mapper.add "tct", "Cycles"
    mapper.add "tsg", "teleopScoring"
    mapper.add "fo", "fedOthers#Pieces"
    mapper.add "of", "othersFed#Pieces"
    mapper.add "dc", "droppedPieces"
    mapper.add "dt", "dockingTimer"
    mapper.add "fs", "finalStatus"
    mapper.add "dn", "totalDockedBots"
    mapper.add "ds", "driverSkill"
    mapper.add "dr", "defenseRating"
    mapper.add "wd", "wasDefended"
    mapper.add "die", "died/immobilized"
    mapper.add "tip", "Tippy?"
    mapper.add "co", "Comments"

    ' 2023 Fields
    ' Auto
    mapper.add "as", "autoStartingLocation"
    mapper.add "asg", "autoScoredGrid"
    mapper.add "acc", "autoCrossedCable"
    mapper.add "acs", "autoCrossedChargingStation"
    mapper.add "am", "autoMobility"
    mapper.add "ad", "autoDocked"
    
    ' Teleop
    mapper.add "tct", "cycleTimes"
    mapper.add "tsg", "scoredGrid"
    mapper.add "tfc", "feedCount"
    mapper.add "wf", "wasFed"
    mapper.add "wd", "wasDefended"
    mapper.add "who", "whoDefended"
    mapper.add "lnk", "smartLinks"
    mapper.add "fpu", "floorPickUp"
    mapper.add "dt", "dockingTime"
    mapper.add "fs", "finalState"
    mapper.add "dn", "numOfRobotsDocked"
    
    'Endgame
    mapper.add "ds", "driverSkill"
    mapper.add "ls", "linksScored"
    mapper.add "dr", "defenseRating"
    mapper.add "sd", "swerveDrive"
    mapper.add "sr", "speedRating"
    mapper.add "die", "diedOrTipped"
    mapper.add "tip", "tippy"
    mapper.add "dc", "droppedCones"
    mapper.add "all", "goodPartner"
    mapper.add "co", "comments"

    If inp = "Camera" Then
        Exit Sub
    End If

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
            key = par(0)
            value = par(1)
            If mapper.Exists(key) Then
                key = mapper(key)
            End If
            data.add key, value
        Next

        tableexists = False
        
        Dim tbl As ListObject
        Dim sht As Worksheet

        'Loop through each sheet and table in the workbook
        For Each sht In ThisWorkbook.Worksheets
            For Each tbl In sht.ListObjects
                If tbl.name = tableName Then
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
            ws.ListObjects.add(xlSrcRange, Range("A1:CZ1"), , xlYes).name = tableName
            i = 0
            Set table = ws.ListObjects(tableName)
            For Each key In data.Keys
                table.Range(i + 1) = key
                i = i + 1
            Next
        End If
        
        Dim newrow As ListRow
        
        Set newrow = table.ListRows.Add
            
        For Each str In data.Keys
            ' Specific data manipulation
            If str = "autoStartingLocation" Then
                data(str) = stripShootingLocation(data(str))
            End If

            newrow.Range(table.ListColumns(str).Index) = data(str)
        Next
    End If
End Sub
Sub savePitData(inp As String)
    Dim fields
    Dim par
    Dim value
    Dim key
    Dim table As ListObject
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim mapper
    Set mapper = CreateObject("Scripting.Dictionary")
    Dim data
    Set data = CreateObject("Scripting.Dictionary")
    Dim tableName As String
    tableName = "PitData"

    ' Set up map
    mapper.add "t", "teamNumber"
    mapper.add "wid", "width"
    mapper.add "wei", "weight"
    mapper.add "drv", "drivetrain"
    mapper.add "odt", "otherDrivetrain"
    mapper.add "sr", "swerveRatio"
    mapper.add "mot", "drivetrainMotor"
    mapper.add "fco", "floorPickUpCones"
    mapper.add "fcu", "floorPickUpCubes"
    mapper.add "ccs", "crossCS"
    mapper.add "aut", "autos"
    
    If inp = "Camera" Then
        Exit Sub
    End If

    If inp = "" Then
        Exit Sub
    End If

    ' MsgBox (inp)
    
    fields = Split(inp, ";")
    If ArrayLen(fields) > 0 Then
        Dim i As Integer
        Dim str

        i = 0

        For Each str In fields
            par = Split(str, "=")
            key = par(0)
            value = par(1)
            If mapper.Exists(key) Then
                key = mapper(key)
            End If
            data.add key, value
        Next

        tableexists = False
        
        Dim tbl As ListObject
        Dim sht As Worksheet

        'Loop through each sheet and table in the workbook
        For Each sht In ThisWorkbook.Worksheets
            For Each tbl In sht.ListObjects
                If tbl.name = tableName Then
                    tableexists = True
                    Set table = tbl
                    Set ws = sht
                End If
            Next tbl
        Next sht
        
        If tableexists Then
            ' Set table = ws.ListObjects(tableName)
        Else
            Dim tablerange As Range
            ws.ListObjects.add(xlSrcRange, Range("A1:CZ1"), , xlYes).name = tableName
            i = 0
            Set table = ws.ListObjects(tableName)
            For Each key In data.Keys
                table.Range(i + 1) = key
                i = i + 1
            Next
        End If

        Dim newrow As ListRow
        
        Set newrow = table.ListRows.add
                
        For Each str In data.Keys
            newrow.Range(table.ListColumns(str).Index) = data(str)
        Next
    End If
End Sub
