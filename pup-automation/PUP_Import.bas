Attribute VB_Name = "PUP_Import"
Option Explicit

' Importerar en DRE-rapport i PUP1 och uppdaterar bladet PUP i ett steg:
'   1. Valj DRE-filen (xlsx/xls/xlsm/csv)
'   2. Track Date / Track Time sätts automatiskt till nu
'      Bara rader med Origin Cntry = SE och Dest Cntry inte SE tas med
'   3. Tabellen Table1 pa bladet Data tomms och fylls (A-G), tabellen Table2 (J-O) far datum och tid
'   4. Refresh All - PUP-bladet raknas om av de befintliga Power Query-fragorna
'   5. Track Time fylls ut till 4 tecken (0905 i stallet for 905), Dropbox Number sätts i kolumn T och bladet PUP visas

Private Const DATA_SHEET As String = "Data"
Private Const PUP_SHEET As String = "PUP"

' PUP1-arbetsboken som uppdateras (den aktiva när makrot startas)
Private wb As Workbook

Public Sub ImporteraDRE()
    Dim fields As Variant
    fields = Array("Tracking Number", "Form ID", "Dest City Name", "Dest Cntry", _
                   "Dest Postal Cd", "Base Service", "Pkg Type")

    Dim wsData As Worksheet, loData As ListObject, loParam As ListObject
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsData = wb.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If Not wsData Is Nothing Then
        Set loData = FindTable(wsData, "Tracking Number")
        Set loParam = FindTable(wsData, "LocID")
    End If
    If loData Is Nothing Or loParam Is Nothing Then
        MsgBox "Den aktiva arbetsboken (" & wb.Name & ") är inte PUP1." & vbLf & vbLf & _
               "Öppna PUP1, klicka på dess fönster och kör makrot igen.", vbExclamation
        Exit Sub
    End If

    ' --- Vilken fil, vilket datum och vilken tid ------------------------------
    Dim path As Variant
    On Error Resume Next
    ChDrive Environ$("USERPROFILE")
    ChDir Environ$("USERPROFILE") & "\Downloads"
    On Error GoTo 0
    path = Application.GetOpenFilename("DRE-rapport (*.xlsx;*.xls;*.xlsm;*.csv),*.xlsx;*.xls;*.xlsm;*.csv", , "Välj DRE-rapporten")
    If VarType(path) = vbBoolean Then Exit Sub

    Dim d As String, t As String
    d = Format$(Now, "MMdd")   ' dagens datum och tid, ingen fråga
    t = Format$(Now, "HHmm")
    If Not (d Like "####") Or Not (t Like "####") Then
        MsgBox "Datum och tid ska vara fyra siffror, t.ex. 0521 och 0905.", vbExclamation
        Exit Sub
    End If

    ' --- Samma krav som Validate Scans ställer, så att du slipper FAIL på alla rader ---
    Dim dt As Date, diff As Long, stopVal As String, db As String
    dt = DateSerial(Year(Now), CLng(Left$(d, 2)), CLng(Right$(d, 2)))
    If Month(dt) <> CLng(Left$(d, 2)) Or Day(dt) <> CLng(Right$(d, 2)) Then
        MsgBox "Track Date " & d & " är inte ett giltigt datum.", vbExclamation
        Exit Sub
    End If
    diff = DateDiff("d", Now, dt)
    If diff >= 7 Or diff <= -7 Then
        If MsgBox("Track Date " & d & " ligger mer än 7 dagar från idag. Valideringen i PUP1 ger FAIL på det." & vbLf & vbLf & "Fortsätta ändå?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    stopVal = UCase$(ParamText(loParam, "Stop Type"))
    If stopVal = "R" Or stopVal = "D" Then
        db = Trim$(CStr(wb.Worksheets(PUP_SHEET).Range("T3").Value2))
        If Not (db Like "#*") Then db = ""
        db = InputBox("Dropbox Number (valfritt, 1-7 siffror). Lämna tomt för att hoppa över:", "PUP", db)
        If Len(db) = 0 Then
            ' hoppas över - Validate Scans ger FAIL, men Send Scans kontrollerar inte det
        ElseIf Len(db) > 7 Or Not (db Like String(Len(db), "#")) Then
            MsgBox "Dropbox Number ska vara 1-7 siffror.", vbExclamation
            Exit Sub
        End If
    End If

    Dim src As Workbook, wsSrc As Worksheet
    Dim headerRow As Long, lastRow As Long, r As Long, i As Long, n As Long, rowsMax As Long
    Dim srcCols() As Long, missing As String, out() As Variant, v As Variant
    Dim vals() As Variant, cn As WorkbookConnection, msg As String
    Dim loPup As ListObject, rngT As Range, rngS As Range, cl As Range
    Dim s As String, lim As Long, noSvc As Long
    Dim colOrigin As Long, skipOrigin As Long, skipDest As Long

    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' --- Läs DRE-filen ------------------------------------------------------
    Set src = Workbooks.Open(Filename:=CStr(path), ReadOnly:=True, UpdateLinks:=0)
    Set wsSrc = src.Worksheets(1)

    headerRow = FindHeaderRow(wsSrc)
    If headerRow = 0 Then
        msg = "Hittar ingen rubrikrad med 'Tracking Number' i DRE-filen."
        GoTo Abort
    End If

    ReDim srcCols(0 To UBound(fields))
    For i = 0 To UBound(fields)
        srcCols(i) = FindColumn(wsSrc, headerRow, CStr(fields(i)))
        If srcCols(i) = 0 Then missing = missing & vbLf & "  - " & fields(i)
    Next i
    If Len(missing) > 0 Then
        msg = "Följande kolumner saknas i DRE-filen:" & missing
        GoTo Abort
    End If

    colOrigin = FindColumn(wsSrc, headerRow, "Origin Cntry")
    If colOrigin = 0 Then
        msg = "Kolumnen 'Origin Cntry' saknas i DRE-filen. Lägg till den bland rapportens fält i DRE."
        GoTo Abort
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCols(0)).End(xlUp).Row
    rowsMax = lastRow - headerRow
    If rowsMax < 1 Then rowsMax = 1
    ReDim out(1 To rowsMax, 0 To UBound(fields))
    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(wsSrc.Cells(r, srcCols(0)).Value2))) > 0 Then
            If UCase$(Trim$(CStr(wsSrc.Cells(r, colOrigin).Value2))) <> "SE" Then
                skipOrigin = skipOrigin + 1               ' ursprungsland är inte SE
            ElseIf UCase$(Trim$(CStr(wsSrc.Cells(r, srcCols(3)).Value2))) = "SE" Then
                skipDest = skipDest + 1                   ' destinationsland är SE
            Else
                n = n + 1
                For i = 0 To UBound(fields)
                    v = wsSrc.Cells(r, srcCols(i)).Value2
                    If VarType(v) = vbString Then v = Trim$(v)   ' siffror lämnas som siffror, som vid kopiera-klistra
                    out(n, i) = v
                Next i
            End If
        End If
    Next r
    src.Close SaveChanges:=False
    Set src = Nothing

    If n = 0 Then
        msg = "Inga rader kvar efter filtret (Origin Cntry = SE, Dest Cntry inte SE)." & vbLf & skipOrigin & " rader hade annat ursprungsland, " & skipDest & " rader gick till SE."
        GoTo Abort
    End If

    ' --- Data-bladet: Table1 (A-G) och Table2 (datum/tid) ---------------------
    If Not loData.DataBodyRange Is Nothing Then loData.DataBodyRange.ClearContents
    loData.Resize loData.HeaderRowRange.Resize(n + 1)

    For i = 0 To UBound(fields)
        ReDim vals(1 To n, 1 To 1)
        For r = 1 To n
            vals(r, 1) = out(r, i)
        Next r
        loData.ListColumns(CStr(fields(i))).DataBodyRange.Value2 = vals
    Next i

    SetParam loParam, "Track Date", CLng(d)
    SetParam loParam, "Track Time", CLng(t)

    ' --- Uppdatera (PUP-bladet byggs av Power Query) ----------------------------
    On Error Resume Next
    For Each cn In wb.Connections
        cn.OLEDBConnection.BackgroundQuery = False
    Next cn
    On Error GoTo Fail
    wb.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    ' --- Efterarbete på PUP-bladet ----------------------------------------------
    Set loPup = FindTable(wb.Worksheets(PUP_SHEET), "AWB Number")
    If Not loPup Is Nothing Then
        If Not loPup.DataBodyRange Is Nothing Then
            lim = loPup.DataBodyRange.Rows.Count
            If lim > n Then lim = n

            ' Power Query gör Track Time till tal, så 0905 blir 905 - fyll ut till 4 tecken
            Set rngT = ColumnRange(loPup, "Track Time", lim)
            If Not rngT Is Nothing Then
                For Each cl In rngT.Cells
                    s = CStr(cl.Value2)
                    If Len(s) > 0 And Len(s) < 4 Then
                        If s Like String(Len(s), "#") Then
                            cl.NumberFormat = "@"
                            cl.Value2 = Right$("0000" & s, 4)
                        End If
                    End If
                Next cl
            End If

            Set rngS = ColumnRange(loPup, "Base Service", lim)
            If Not rngS Is Nothing Then
                For Each cl In rngS.Cells
                    If Not IsError(cl.Value2) Then
                        If Len(CStr(cl.Value2)) = 0 Then noSvc = noSvc + 1
                    End If
                Next cl
            End If
        End If
    End If

    ' Dropbox Number (kolumn T) ligger utanför frågan - fyll alla rader, rensa gamla värden
    With wb.Worksheets(PUP_SHEET)
        .Range("T3:T1000").ClearContents
        If Len(db) > 0 Then
            .Range("T3").Resize(n, 1).NumberFormat = "@"
            .Range("T3").Resize(n, 1).Value2 = db
        End If
    End With

    Application.ScreenUpdating = True
    On Error Resume Next
    wb.Activate
    wb.Worksheets(PUP_SHEET).Activate
    On Error GoTo 0

    msg = n & " försändelser importerade och PUP är uppdaterad. Kör nu Validate Scans."
    If skipOrigin + skipDest > 0 Then msg = msg & vbLf & vbLf & "Hoppade över " & skipOrigin + skipDest & " rader: " & skipOrigin & " med annat ursprungsland än SE, " & skipDest & " med Dest Cntry SE."
    If n > 998 Then msg = msg & vbLf & vbLf & "OBS: PUP1:s egna knappar rensar och validerar bara rad 3-1000 (998 rader)."
    If noSvc > 0 Then msg = msg & vbLf & vbLf & noSvc & " rader har ingen Base Service (värdet finns inte i Service-tabellen)."
    MsgBox msg, vbInformation
    If MsgBox("Vill du köra Validate Scans och Send Scans nu?", vbYesNo + vbQuestion) = vbYes Then ValidateAndSend n
    Exit Sub

Abort:
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox msg, vbCritical
    Exit Sub

Fail:
    msg = "Något gick fel: " & Err.Description
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox msg, vbCritical
End Sub

' Kör PUP1:s egna Validate Scans och Send Scans (samma som knapparna). Inloggningen
' (användarnamn, lösenord, Device ID) skriver du själv i PUP1:s ruta - den sparas inte här.
Private Sub ValidateAndSend(n As Long)
    Dim ws As Worksheet, fails As Long, q As String
    Set ws = wb.Worksheets(PUP_SHEET)
    wb.Activate
    ws.Activate

    On Error GoTo Oops
    Application.Run "'" & wb.Name & "'!PUP_ValidateScans_Button_Click"

    fails = Application.WorksheetFunction.CountIf(ws.Range("W3:W1000"), "FAIL")
    If fails > 0 Then
        q = fails & " rader har FAIL i valideringen. Skicka ändå?"
    Else
        q = "Valideringen är godkänd. Skicka " & n & " scans?"
    End If
    q = q & vbLf & vbLf & "Du får logga in i nästa ruta."
    If MsgBox(q, vbYesNo + vbExclamation + IIf(fails > 0, vbDefaultButton2, 0)) = vbNo Then Exit Sub

    ws.Activate
    Application.Run "'" & wb.Name & "'!PUP_SendScans_Button_Click"
    Exit Sub

Oops:
    MsgBox "Något gick fel i PUP1:s makro: " & Err.Description, vbCritical
End Sub

' Tabellen på bladet vars första kolumn börjar med headerStart (versaler, mellanslag och radbrytningar ignoreras).
Private Function FindTable(ws As Worksheet, headerStart As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If Norm(lo.ListColumns(1).Name) Like Norm(headerStart) & "*" Then
            Set FindTable = lo
            Exit Function
        End If
    Next lo
End Function

' Första dataraden i kolumnen vars rubrik börjar med headerStart.
Private Sub SetParam(lo As ListObject, headerStart As String, v As Variant)
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If Norm(c.Name) Like Norm(headerStart) & "*" Then
            With c.DataBodyRange.Cells(1, 1)
                .NumberFormat = "0000"
                .Value2 = v
            End With
            Exit Sub
        End If
    Next c
End Sub

' Texten i första dataraden i kolumnen vars rubrik börjar med headerStart.
Private Function ParamText(lo As ListObject, headerStart As String) As String
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If Norm(c.Name) Like Norm(headerStart) & "*" Then
            ParamText = Trim$(CStr(c.DataBodyRange.Cells(1, 1).Value2))
            Exit Function
        End If
    Next c
End Function

' De första cnt raderna i kolumnen vars rubrik börjar med headerStart.
Private Function ColumnRange(lo As ListObject, headerStart As String, cnt As Long) As Range
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If Norm(c.Name) Like Norm(headerStart) & "*" Then
            Set ColumnRange = c.DataBodyRange.Resize(cnt)
            Exit Function
        End If
    Next c
End Function

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long
    For r = 1 To 30
        For c = 1 To 60
            If IsAlias("Tracking Number", CStr(ws.Cells(r, c).Value2)) Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindColumn(ws As Worksheet, headerRow As Long, field As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If IsAlias(field, CStr(ws.Cells(headerRow, c).Value2)) Then
            FindColumn = c
            Exit Function
        End If
    Next c
End Function

' Godtar fältnamnet och vanliga varianter.
Private Function IsAlias(field As String, header As String) As Boolean
    Dim h As String, a As Variant, x As Variant
    h = Norm(header)
    If Len(h) = 0 Then Exit Function
    Select Case field
        Case "Tracking Number": a = Array("trackingnumber", "trackingnbr", "trackingno", "trknbr", "tracknbr", "trackingid", "awbnumber")
        Case "Form ID":         a = Array("formid", "formcode", "formcd")
        Case "Dest City Name":  a = Array("destcityname", "destcity", "destinationcity", "destinationcityname")
        Case "Dest Cntry":      a = Array("destcntry", "destcountry", "destcntrycd", "destcountrycode", "destinationcountry")
        Case "Dest Postal Cd":  a = Array("destpostalcd", "destpostalcode", "destinationpostalcode", "destzip", "destpostal")
        Case "Base Service":    a = Array("baseservice", "baseservicecd", "service")
        Case "Pkg Type":        a = Array("pkgtype", "packagetype", "pkgtypecd")
        Case "Origin Cntry":    a = Array("origincntry", "origincountry", "origcntry", "origcountry", "origincntrycd", "origincountrycode", "shippercountry", "shippercntry")
        Case Else:              a = Array(Norm(field))
    End Select
    For Each x In a
        If h = x Then
            IsAlias = True
            Exit Function
        End If
    Next x
End Function

Private Function Norm(s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = LCase$(Mid$(s, i, 1))
        If ch Like "[a-z0-9]" Then Norm = Norm & ch
    Next i
End Function
