Attribute VB_Name = "PUP_Import"
Option Explicit

' Importerar en DRE-export till PUP1-arbetsboken i ett steg:
'   1. Fragar efter DRE-filen (xlsx/xls/csv)
'   2. Hittar kolumnerna Tracking Number ... Pkg Type via rubrikerna
'   3. Tommer tabellen pa bladet "Data" och fyller den med exakt sa manga rader som behovs
'   4. Satter Track Date (MMdd) och Track Time (HHmm) i parametertabellen (J-O)
'   5. Uppdaterar allt (Data -> Refresh All) och visar bladet "PUP"

Private Const DATA_SHEET As String = "Data"
Private Const PUP_SHEET As String = "PUP"

Public Sub ImporteraDRE()
    Dim fields As Variant
    fields = Array("Tracking Number", "Form ID", "Dest City Name", "Dest Cntry", _
                   "Dest Postal Cd", "Base Service", "Pkg Type")

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)

    Dim loData As ListObject, loParam As ListObject
    Set loData = FindTable(wsData, "Tracking Number")
    Set loParam = FindTable(wsData, "LocID")
    If loData Is Nothing Then
        MsgBox "Hittar ingen tabell med kolumnen 'Tracking Number' pa bladet '" & DATA_SHEET & "'.", vbCritical
        Exit Sub
    End If

    Dim path As Variant
    path = Application.GetOpenFilename("DRE-export (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", , "Valj DRE-exporten")
    If VarType(path) = vbBoolean Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' --- Las DRE-filen -------------------------------------------------------
    Dim src As Workbook
    Set src = Workbooks.Open(Filename:=CStr(path), ReadOnly:=True, UpdateLinks:=0)
    Dim wsSrc As Worksheet
    Set wsSrc = src.Worksheets(1)

    Dim headerRow As Long, srcCols() As Long, missing As String
    headerRow = FindHeaderRow(wsSrc)
    If headerRow = 0 Then
        src.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Hittar ingen rubrikrad med 'Tracking Number' i DRE-filen.", vbCritical
        Exit Sub
    End If

    ReDim srcCols(LBound(fields) To UBound(fields))
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        srcCols(i) = FindColumn(wsSrc, headerRow, CStr(fields(i)))
        If srcCols(i) = 0 Then missing = missing & vbLf & "  - " & fields(i)
    Next i
    If Len(missing) > 0 Then
        src.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Foljande kolumner saknas i DRE-filen:" & missing, vbCritical
        Exit Sub
    End If

    Dim lastRow As Long, r As Long, n As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCols(0)).End(xlUp).Row
    Dim out() As Variant
    ReDim out(1 To Application.Max(1, lastRow - headerRow), 0 To UBound(fields))
    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(wsSrc.Cells(r, srcCols(0)).Value))) > 0 Then
            n = n + 1
            For i = 0 To UBound(fields)
                out(n, i) = Trim$(CStr(wsSrc.Cells(r, srcCols(i)).Value))
            Next i
        End If
    Next r
    src.Close SaveChanges:=False

    If n = 0 Then
        Application.ScreenUpdating = True
        MsgBox "DRE-filen innehaller inga rader.", vbExclamation
        Exit Sub
    End If

    ' --- Skriv till tabellen pa Data ----------------------------------------
    If Not loData.DataBodyRange Is Nothing Then loData.DataBodyRange.ClearContents
    loData.Resize loData.HeaderRowRange.Resize(n + 1)

    Dim col As ListColumn, vals() As Variant
    For i = 0 To UBound(fields)
        Set col = loData.ListColumns(CStr(fields(i)))
        ReDim vals(1 To n, 1 To 1)
        For r = 1 To n
            vals(r, 1) = out(r, i)
        Next r
        ' Text sa att tracking- och postnummer inte tappar nollor eller blir 7,9E+11
        col.DataBodyRange.NumberFormat = "@"
        col.DataBodyRange.Value = vals
    Next i

    ' --- Parametrar J-O -----------------------------------------------------
    If Not loParam Is Nothing Then
        Application.ScreenUpdating = True
        Dim d As String, t As String
        d = InputBox("Track Date (MMdd):", "PUP", Format$(Now, "MMdd"))
        If Len(d) > 0 Then SetParam loParam, "Track Date", d
        t = InputBox("Track Time (HHmm):", "PUP", Format$(Now, "HHmm"))
        If Len(t) > 0 Then SetParam loParam, "Track Time", t
        Application.ScreenUpdating = False
    End If

    ' --- Uppdatera PUP ------------------------------------------------------
    Dim cn As WorkbookConnection
    For Each cn In ThisWorkbook.Connections
        On Error Resume Next
        cn.OLEDBConnection.BackgroundQuery = False
        On Error GoTo Fail
    Next cn
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Worksheets(PUP_SHEET).Activate
    On Error GoTo 0
    MsgBox n & " forsandelser importerade och PUP ar uppdaterad.", vbInformation
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Nagot gick fel: " & Err.Description, vbCritical
End Sub

' Tabellen pa bladet som har en kolumn vars rubrik borjar med headerStart.
Private Function FindTable(ws As Worksheet, headerStart As String) As ListObject
    Dim lo As ListObject, c As ListColumn
    For Each lo In ws.ListObjects
        For Each c In lo.ListColumns
            If Norm(c.Name) Like Norm(headerStart) & "*" Then
                Set FindTable = lo
                Exit Function
            End If
        Next c
    Next lo
End Function

' Satter forsta dataraden i den kolumn vars rubrik borjar med headerStart.
Private Sub SetParam(lo As ListObject, headerStart As String, value As String)
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If Norm(c.Name) Like Norm(headerStart) & "*" Then
            c.DataBodyRange.Cells(1, 1).NumberFormat = "@"
            c.DataBodyRange.Cells(1, 1).Value = value
            Exit Sub
        End If
    Next c
End Sub

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long
    For r = 1 To 30
        For c = 1 To 60
            If IsAlias("Tracking Number", CStr(ws.Cells(r, c).Value)) Then
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
        If IsAlias(field, CStr(ws.Cells(headerRow, c).Value)) Then
            FindColumn = c
            Exit Function
        End If
    Next c
End Function

' Godtar faltnamnet och vanliga varianter (versaler, mellanslag och punkter ignoreras).
Private Function IsAlias(field As String, header As String) As Boolean
    Dim h As String, a As Variant
    h = Norm(header)
    If Len(h) = 0 Then Exit Function
    Select Case field
        Case "Tracking Number": a = Array("trackingnumber", "trackingnbr", "trackingno", "trknbr", "tracknbr", "trackingid")
        Case "Form ID":         a = Array("formid", "formcode", "formcd")
        Case "Dest City Name":  a = Array("destcityname", "destcity", "destinationcity", "destinationcityname")
        Case "Dest Cntry":      a = Array("destcntry", "destcountry", "destcntrycd", "destcountrycode", "destinationcountry")
        Case "Dest Postal Cd":  a = Array("destpostalcd", "destpostalcode", "destinationpostalcode", "destzip", "destpostal")
        Case "Base Service":    a = Array("baseservice", "baseservicecd", "service")
        Case "Pkg Type":        a = Array("pkgtype", "packagetype", "pkgtypecd")
        Case Else:              a = Array(Norm(field))
    End Select
    Dim x As Variant
    For Each x In a
        If h = x Then IsAlias = True: Exit Function
    Next x
End Function

Private Function Norm(s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = LCase$(Mid$(s, i, 1))
        If ch Like "[a-z0-9]" Then Norm = Norm & ch
    Next i
End Function
