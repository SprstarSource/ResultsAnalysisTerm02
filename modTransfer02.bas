Attribute VB_Name = "modTransfer02"
Option Explicit

Private m_runSucceeded As Boolean
Private m_lastReport As String
Private m_silent As Boolean

' ============================================================
' Transfer 02: \Source02\ -> Term02Guidance.xlsx (Data science; Data01 arts)
' Target file: inside the output folder next to the tool workbook (created when missing).
' If the file is only found beside the tool (old location) it is copied into the folder first.
' ============================================================

Private Const SRC_FOLDER As String = "Source02"
Private Const DST_FILE As String = "Term02Guidance.xlsx"
Private Const DST_SHEET As String = "Data"
Private Const DST_TABLE As String = "Data"
Private Const HDR_ROW As Long = 2
Private Const FIRST_ROW As Long = 3
Private Const NCOLS As Long = 21          ' target columns A..U (U already exists)
Private Const COL_SECTION As Long = 14    ' N
Private Const COL_WISH1 As Long = 15      ' O
Private Const COL_WISH2 As Long = 16      ' P
Private Const SRC_LASTCOL As Long = 33    ' science A..AG; arts A..AE
Private Const GROUP_SIZE As Long = 2      ' adjacent first/second term marks, as modAverage02

' Sheet title prefix stored as hex code points (keeps this file pure ASCII)
Private Const TITLE_CODES As String = "623 648 644 649 20 62B 627 646 648 64A 20 62C 630 639 20 645 634 62A 631 643 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627"

' ---- Collection engine constants ----
Private Const DATA_START_ROW       As Long = 6      ' header row in the source sheet
Private Const TRAILING_ROWS_SKIP   As Long = 1      ' footer rows after last used row
Private Const IDENTITY_LAST_COL    As Long = 5      ' columns A:E = identity columns (AutoFit)
Private Const COL_WIDTH            As Double = 5.7  ' width for columns F:LastCol
Private Const FONT_SIZE            As Long = 14
Private Const FONT_ARABIC          As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC         As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK        As String = "Calibri"

Private Const PREFERRED_FONT As String = "Cairo"
Private Const FALLBACK_FONT  As String = "Arial"
Private Const HEADER_FONT_SIZE As Double = 11
Private Const HEADER_ROW_HEIGHT As Double = 30

Private g_ActiveArabicFont  As String
Private g_ActiveNumericFont As String

Private Type TransferBatch
    Data As Variant
    Formats As Variant
    Rows As Long
    SheetCount As Long
    IsArts As Boolean
    Skipped As String
    ErrorCells As Long
    NonNumericCells As Long
End Type

' ---- واجهات التحكم في الوضع الصامت والتقرير ----
Public Sub SetSilentMode(ByVal silent As Boolean)
    m_silent = silent
End Sub

Public Function LastReport() As String
    LastReport = m_lastReport
End Function

Public Sub Run_Transfer02_Silent()
    m_silent = True
    Run_Transfer02
    m_silent = False
End Sub

Public Function Transfer02Succeeded() As Boolean
    Transfer02Succeeded = m_runSucceeded
End Function

Public Function Transfer02Report() As String
    Transfer02Report = m_lastReport
End Function

Public Sub Run_Transfer02()
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldEvents As Boolean, stateSaved As Boolean
    Dim wasSilent As Boolean
    Dim scienceTouched As Boolean, artsTouched As Boolean
    Dim wbSrc As Workbook, wbDst As Workbook, srcOpened As Boolean, dstOpened As Boolean
    Dim science As TransferBatch, arts As TransferBatch, problems As Object
    Dim ws As Worksheet, lo As ListObject, failed As Boolean, errText As String
    Dim srcCopied As Long, srcSkipped As Long, srcList As String

    wasSilent = m_silent
    m_runSucceeded = False
    m_lastReport = ""
    On Error GoTo EH
    Application.EnableCancelKey = xlErrorHandler

    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    stateSaved = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 1, , "Save the tool workbook before running this procedure."
    If Not SrcBuildWorkbook(ThisWorkbook.Path & "\" & SRC_FOLDER & "\", wbSrc, srcCopied, srcSkipped, srcList) Then
        Err.Raise vbObjectError + 4, , "No usable source file was found in " & SRC_FOLDER & "." & vbLf & srcList
    End If
    srcOpened = True
    Set wbDst = OpenWb(ResolveDstPath(), False, dstOpened)
    If wbDst.ReadOnly Then Err.Raise vbObjectError + 2, , DST_FILE & " is open read-only."
    Set problems = CreateObject("Scripting.Dictionary")
    PrepareTransfer wbSrc, wbDst, False, science, problems
    PrepareTransfer wbSrc, wbDst, True, arts, problems
    RaiseValidationIssues problems, "Header or data errors: sheet, cell/row and expected column are listed."
    If science.SheetCount + arts.SheetCount = 0 Then Err.Raise vbObjectError + 3, , "No matching science or arts sheet found. Target left untouched."

    If science.SheetCount > 0 Then Set lo = GuidanceTable(RequireWorksheet(wbDst, "Data"), 21)
    If arts.SheetCount > 0 Then
        artsTouched = True
        EnsureArtsTarget wbDst
        Set lo = GuidanceTable(RequireWorksheet(wbDst, "Data01"), 15)
    End If
    If science.SheetCount > 0 Then
        scienceTouched = True
        CheckTrackWrite wbDst, RequireWorksheet(wbDst, "Data")
    End If
    If arts.SheetCount > 0 Then CheckTrackWrite wbDst, RequireWorksheet(wbDst, "Data01")
    If science.SheetCount > 0 Then WriteTransfer wbDst, science
    If arts.SheetCount > 0 Then WriteTransfer wbDst, arts
    wbDst.Save
    m_runSucceeded = True
    m_lastReport = TransferBatchMessage(science) & vbLf & TransferBatchMessage(arts) & science.Skipped & arts.Skipped

Done:
    On Error Resume Next
    If srcOpened Then wbSrc.Close SaveChanges:=False
    If failed And dstOpened Then
        wbDst.Close SaveChanges:=False
    ElseIf failed And Not wbDst Is Nothing Then
        For Each ws In wbDst.Worksheets
            If (ws.Name = "Data" And scienceTouched) Or (ws.Name = "Data01" And artsTouched) Then
                Set lo = Nothing
                If ws.Name = "Data01" Then Set lo = GuidanceTable(ws, 15) Else Set lo = GuidanceTable(ws, 21)
                If Not lo Is Nothing Then LockAndProtect ws, lo.DataBodyRange
            End If
        Next ws
    End If
    If stateSaved Then
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
    End If
    Application.EnableCancelKey = xlInterrupt
    m_silent = False
    On Error GoTo 0
    If failed Then
        m_lastReport = "Transfer failed:" & vbLf & errText
        If Not wasSilent Then MsgBox m_lastReport, vbCritical, "Transfer 02"
    Else
        If Not wasSilent Then
            MsgBox m_lastReport, IIf(science.ErrorCells + science.NonNumericCells + arts.ErrorCells + arts.NonNumericCells > 0, vbExclamation, vbInformation), "Transfer 02"
        End If
    End If
    Exit Sub

EH:
    m_runSucceeded = False
    failed = True
    If Err.Number = 18 Then
        errText = "Operation interrupted by user."
    Else
        errText = Err.Description
        If (scienceTouched Or artsTouched) And Not dstOpened Then
            errText = errText & vbLf & vbLf & ArtsText("failure_unsaved")
        End If
    End If
    Resume Done
End Sub

Private Sub PrepareTransfer(ByVal wbSrc As Workbook, ByVal wbDst As Workbook, ByVal isArts As Boolean, _
                            ByRef batch As TransferBatch, ByVal problems As Object)
    Dim ws As Worksheet, prefix As String, h As String, rest As String
    Dim idx() As Long, groups() As Long, rows() As Long, m As Long, total As Long
    Dim map As Variant, arr As Variant, output() As Variant, formats() As Variant
    Dim j As Long, i As Long, c As Long, p As Long, width As Long, copied As Long, lastCol As Long, lastRow As Long
    Dim born As Date, location As String
    batch.IsArts = isArts
    If isArts Then
        prefix = Norm(ArtsText("arts_prefix"))
        map = Array(1, 2, 3, 4, 5, 6, 16, 18, 20, 30)
        width = 15: copied = 10: lastCol = 31
    Else
        prefix = TitlePrefix()
        map = Array(1, 2, 3, 4, 5, 6, 8, 10, 12, 16, 22, 24, 32)
        width = 21: copied = 13: lastCol = 33
    End If
    For Each ws In wbSrc.Worksheets
        ws.Range("H1").Calculate
        h = Norm(ws.Range("H1").MergeArea.Cells(1, 1).Value2)
        h = Replace(h, ChrW(&H62C) & ChrW(&H62F) & ChrW(&H639), ChrW(&H62C) & ChrW(&H630) & ChrW(&H639))
        If Len(h) >= Len(prefix) And Len(prefix) > 0 Then
            If Left$(h, Len(prefix)) = prefix Then
                rest = Trim$(Mid$(h, Len(prefix) + 1))
                If PositiveGroup(rest) Or (isArts And Len(rest) = 0) Then
                    ValidateSourceHeaders ws, isArts, map, problems
                    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
                    If lastRow >= FIRST_ROW Then
                        m = m + 1
                        ReDim Preserve idx(1 To m)
                        ReDim Preserve groups(1 To m)
                        ReDim Preserve rows(1 To m)
                        idx(m) = ws.Index
                        If Len(rest) > 0 Then groups(m) = CLng(rest)
                        rows(m) = lastRow - FIRST_ROW + 1
                        total = total + rows(m)
                    Else
                        batch.Skipped = batch.Skipped & vbLf & ws.Name & " (no data rows)"
                    End If
                Else
                    batch.Skipped = batch.Skipped & vbLf & ws.Name & " (group number unreadable in H1)"
                End If
            End If
        End If
    Next ws
    batch.Rows = total
    batch.SheetCount = m
    If total = 0 Then Exit Sub
    ReDim output(1 To total, 1 To width)
    ReDim formats(1 To width)
    For c = 1 To width
        formats(c) = "General"
    Next c
    For j = 1 To m
        Set ws = wbSrc.Worksheets(idx(j))
        lastRow = FIRST_ROW + rows(j) - 1
        ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastRow, 5)).Calculate
        For c = 6 To copied
            ws.Range(ws.Cells(FIRST_ROW, CLng(map(c - 1))), ws.Cells(lastRow, CLng(map(c - 1)) + GROUP_SIZE - 1)).Calculate
        Next c
        If j = 1 Then
            For c = 1 To 5
                formats(c) = ws.Cells(FIRST_ROW, CLng(map(c - 1))).NumberFormat
            Next c
            For c = 6 To copied
                formats(c) = "0.00"
            Next c
        End If
        arr = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + rows(j) - 1, lastCol)).Value2
        For i = 1 To rows(j)
            p = p + 1
            For c = 1 To 5
                output(p, c) = arr(i, map(c - 1))
            Next c
            For c = 6 To copied
                output(p, c) = AveragePair(arr, i, CLng(map(c - 1)), batch.ErrorCells, batch.NonNumericCells)
            Next c
            location = ws.Name & "!" & (FIRST_ROW + i - 1)
            If IsError(output(p, 2)) Or BlankValue(output(p, 2)) Then AddValidationIssue problems, location, "B"
            For c = 6 To copied
                If Not IsNum(output(p, c)) Then
                    AddValidationIssue problems, location, ColumnLetter(CLng(map(c - 1))) & ":" & ColumnLetter(CLng(map(c - 1)) + GROUP_SIZE - 1)
                End If
            Next c
            If Not BlankValue(output(p, 3)) Then
                If Not TryBirthDate(output(p, 3), wbSrc.Date1904, born) Then
                    AddValidationIssue problems, location, "C"
                ElseIf VarType(output(p, 3)) = vbDouble Then
                    If wbSrc.Date1904 <> wbDst.Date1904 Then
                        output(p, 3) = CDbl(born)
                        If wbDst.Date1904 Then output(p, 3) = CDbl(output(p, 3)) - 1462
                    End If
                End If
            End If
            If isArts Then
                If IsNum(output(p, 6)) And IsNum(output(p, 7)) And IsNum(output(p, 8)) And IsNum(output(p, 9)) Then
                    output(p, 13) = Rnd2(((output(p, 6) + output(p, 8) + output(p, 9)) * 3 + output(p, 7)) / 10)
                    output(p, 14) = Rnd2((output(p, 6) * 5 + output(p, 8) * 2 + output(p, 9) + output(p, 7) * 2) / 10)
                End If
            Else
                output(p, COL_SECTION) = groups(j)
                FillProfiles output, p
            End If
        Next i
    Next j
    If isArts Then
        formats(13) = "00.00": formats(14) = "00.00"
    Else
        formats(COL_SECTION) = "00"
        For c = 17 To 20
            formats(c) = "00.00"
        Next c
    End If
    batch.Data = output
    batch.Formats = formats
End Sub

Private Sub EnsureArtsTarget(ByVal wb As Workbook)
    Dim ws As Worksheet, anchor As Worksheet, lo As ListObject, made As Boolean, labels As Variant, c As Long, headers(1 To 1, 1 To 15) As Variant
    Set ws = FindWorksheet(wb, "Data01")
    Set anchor = FindWorksheet(wb, "Data")
    If ws Is Nothing Then
        If anchor Is Nothing Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        Else
            Set ws = wb.Worksheets.Add(After:=anchor)
        End If
        ws.Name = "Data01"
        made = True
    ElseIf Not anchor Is Nothing Then
        If ws.Index <> anchor.Index + 1 Then ws.Move After:=anchor
    End If
    ws.DisplayRightToLeft = True
    If ws.ListObjects.Count = 0 Then
        ws.Unprotect Password:=""
        labels = Array("number", "name", "birth", "sex", "repeat", "arabic", "history", "french", "english", "average", "wish1", "wish2", "p_lang", "p_arts", "decision")
        For c = 1 To 15
            headers(1, c) = ArtsText(CStr(labels(c - 1)))
        Next c
        ws.Range("A2:O2").Value2 = headers
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("A2:O2"), XlListObjectHasHeaders:=xlYes)
        lo.Name = "Data01"
        lo.TableStyle = ""
        If made Then
            ws.Range("A1:O1").Merge
            ws.Range("A1").Value2 = ArtsText("arts_title")
            ws.Range("A1:O3").Font.Name = "Dubai"
            ws.Range("A1:O3").Font.Size = 14
            ws.Range("A1:O2").Font.Bold = True
            ws.Range("A1:O3").HorizontalAlignment = xlCenter
            ws.Range("A1:O3").VerticalAlignment = xlCenter
            ws.Columns("A:O").ColumnWidth = 11
            ws.Columns("B").ColumnWidth = 28
            ws.Columns("C").ColumnWidth = 16
            ws.Rows(1).RowHeight = 30
            ws.Rows(2).RowHeight = 120
            ws.Range("A2:O2").Orientation = 90
        End If
    Else
        Set lo = GuidanceTable(ws, 15)
    End If
End Sub

Private Sub WriteTransfer(ByVal wb As Workbook, ByRef batch As TransferBatch)
    Dim ws As Worksheet, lo As ListObject, body As Range, width As Long, c As Long, formats As Variant
    If batch.SheetCount = 0 Or batch.Rows = 0 Then Exit Sub
    If batch.IsArts Then
        Set ws = wb.Worksheets("Data01"): width = 15
    Else
        Set ws = wb.Worksheets("Data"): width = 21
    End If
    Set lo = GuidanceTable(ws, width)
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    InvalidateOrientation wb, ws
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Resize(, width).ClearContents
    lo.Resize ws.Range(ws.Cells(2, 1), ws.Cells(2 + batch.Rows, lo.ListColumns.Count))
    If batch.Rows > 0 Then
        Set body = lo.DataBodyRange
        body.Resize(, width).Value2 = batch.Data
        If batch.IsArts Then
            body.Resize(, width).Font.Name = "Dubai"
            body.Resize(, width).Font.Size = 14
            body.Resize(, width).RowHeight = 20
        End If
        formats = batch.Formats
        For c = 1 To width
            body.Columns(c).NumberFormat = CStr(formats(c))
        Next c
    End If
    LockAndProtect ws, body
End Sub

Private Sub FillProfiles(ByRef a() As Variant, ByVal p As Long)
    Dim ar As Variant, ma As Variant, ph As Variant, ns As Variant
    Dim hg As Variant, inf As Variant, te As Variant
    ar = a(p, 6): ma = a(p, 7): ph = a(p, 8): ns = a(p, 9)
    hg = a(p, 10): inf = a(p, 11): te = a(p, 12)

    If IsNum(ma) And IsNum(ph) And IsNum(ar) Then
        a(p, 17) = Rnd2((ma * 6 + ph * 4 + ar) / 11)
    End If
    If IsNum(te) And IsNum(ma) And IsNum(ph) And IsNum(ar) Then
        a(p, 18) = Rnd2((te * 4 + ma * 3 + ph * 3 + ar) / 11)
    End If
    If IsNum(ns) And IsNum(ma) And IsNum(ph) And IsNum(ar) Then
        a(p, 19) = Rnd2((ns * 4 + ma * 3 + ph * 3 + ar) / 11)
    End If
    If IsNum(hg) And IsNum(ma) And IsNum(inf) And IsNum(ar) Then
        a(p, 20) = Rnd2((hg * 4 + ma * 4 + inf + ar * 2) / 11)
    End If
End Sub

Private Function IsNum(ByVal v As Variant) As Boolean
    If VarType(v) <> vbDouble Then Exit Function
    IsNum = (CDbl(v) >= 0 And CDbl(v) <= 20)
End Function

Private Function Rnd2(ByVal x As Double) As Double
    Rnd2 = Application.WorksheetFunction.Round(x, 2)
End Function

' ==============================================================================
' OUTPUT FOLDER + TARGET FILE PATH
' ==============================================================================
' Name of the output folder next to the tool workbook (hex code points)
Private Function OutputFolderName() As String
    OutputFolderName = Txt("645 639 627 644 62C 629 20 628 64A 627 646 627 62A 20 627 644 641 635 644 20 30 32")
End Function

Private Function FolderExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExistsPath = ((GetAttr(p) And vbDirectory) <> 0)
    If Err.Number <> 0 Then FolderExistsPath = False: Err.Clear
End Function

Private Function EnsureFolder(ByVal p As String) As Boolean
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Not FolderExistsPath(p) Then
        On Error Resume Next
        MkDir p
        Err.Clear
        On Error GoTo 0
    End If
    EnsureFolder = FolderExistsPath(p)
End Function

' Full path of Term02Guidance.xlsx inside the output folder (folder created when missing).
' A copy found beside the tool (old location) is copied into the folder; the original is untouched.
Private Function ResolveDstPath() As String
    Dim outDir As String, dstPath As String, oldPath As String, copyErr As Long
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 1, , "Save the tool workbook before running this procedure."
    outDir = ThisWorkbook.Path & Application.PathSeparator & OutputFolderName()
    If Not EnsureFolder(outDir) Then Err.Raise vbObjectError + 5, , "The output folder could not be created: " & outDir
    dstPath = outDir & Application.PathSeparator & DST_FILE
    If Not FileExistsPath(dstPath) Then
        oldPath = ThisWorkbook.Path & Application.PathSeparator & DST_FILE
        If Not FileExistsPath(oldPath) Then Err.Raise vbObjectError + 1, , "File not found: " & dstPath
        On Error Resume Next
        FileCopy oldPath, dstPath
        copyErr = Err.Number
        Err.Clear
        On Error GoTo 0
        If copyErr <> 0 Then Err.Raise vbObjectError + 6, , "Could not copy " & oldPath & " into the output folder (close it in Excel and retry)."
    End If
    ResolveDstPath = dstPath
End Function

Private Function OpenWb(ByVal fullPath As String, ByVal asReadOnly As Boolean, ByRef opened As Boolean) As Workbook
    Dim wb As Workbook, fileName As String
    opened = False
    fileName = Mid$(fullPath, InStrRev(fullPath, Application.PathSeparator) + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenWb = wb
            Exit Function
        End If
        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1, , "A different file with the same name is open: " & wb.FullName
        End If
    Next wb
    If Not FileExistsPath(fullPath) Then Err.Raise vbObjectError + 1, , "File not found: " & fullPath
    Set OpenWb = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=asReadOnly)
    opened = True
End Function

Private Function TitlePrefix() As String
    Dim parts() As String, i As Long, s As String
    parts = Split(TITLE_CODES, " ")
    For i = 0 To UBound(parts)
        s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    TitlePrefix = Norm(s)
End Function

Private Function Norm(ByVal v As Variant) As String
    Dim t As String, i As Long
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    t = CStr(v)
    t = Replace(t, ChrW(&H623), ChrW(&H627))
    t = Replace(t, ChrW(&H625), ChrW(&H627))
    t = Replace(t, ChrW(&H622), ChrW(&H627))
    t = Replace(t, ChrW(&H649), ChrW(&H64A))
    t = Replace(t, ChrW(&H629), ChrW(&H647))
    t = Replace(t, ChrW(&H640), "")
    t = Replace(t, ChrW(&HA0), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, ChrW(&H200E), "")
    t = Replace(t, ChrW(&H200F), "")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    For i = 0 To 9
        t = Replace(t, ChrW(&H660 + i), CStr(i))
        t = Replace(t, ChrW(&H6F0 + i), CStr(i))
    Next i
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    Norm = Trim$(t)
End Function

Private Function BlankValue(ByVal x As Variant) As Boolean
    If IsError(x) Then Exit Function
    If IsNull(x) Then
        BlankValue = True
    ElseIf IsEmpty(x) Then
        BlankValue = True
    Else
        BlankValue = (Len(Norm(x)) = 0)
    End If
End Function

Private Function DigitsOnly(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    DigitsOnly = True
End Function

Private Function ValidYMD(ByVal ys As String, ByVal ms As String, ByVal ds As String, ByRef d As Date) As Boolean
    Dim y As Long, m As Long, dy As Long
    On Error GoTo InvalidDate
    If Len(ys) <> 4 Or Len(ms) > 2 Or Len(ds) > 2 Then Exit Function
    If Not DigitsOnly(ys) Or Not DigitsOnly(ms) Or Not DigitsOnly(ds) Then Exit Function
    y = CLng(ys): m = CLng(ms): dy = CLng(ds)
    If y < 1950 Or y > Year(Date) Then Exit Function
    If m < 1 Or m > 12 Or dy < 1 Or dy > 31 Then Exit Function
    d = DateSerial(y, m, dy)
    If Year(d) <> y Or Month(d) <> m Or Day(d) <> dy Then Exit Function
    ValidYMD = True
InvalidDate:
End Function

Private Function TryBirthDate(ByVal x As Variant, ByVal date1904 As Boolean, ByRef d As Date) As Boolean
    Dim s As String, parts() As String, serial As Double
    On Error GoTo InvalidDate
    If IsError(x) Then Exit Function
    If BlankValue(x) Then Exit Function
    Select Case VarType(x)
    Case vbDate
        d = DateSerial(Year(x), Month(x), Day(x))
        If Year(d) >= 1950 And Year(d) <= Year(Date) Then TryBirthDate = True
        Exit Function
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        serial = CDbl(x)
        If serial < 1 Or serial >= 100000 Then Exit Function
        serial = Fix(serial)
        If date1904 Then serial = serial + 1462
        d = CDate(serial)
        If Year(d) >= 1950 And Year(d) <= Year(Date) Then TryBirthDate = True
        Exit Function
    Case vbString
        s = Norm(x)
    Case Else
        Exit Function
    End Select
    s = Replace(s, "-", "/")
    s = Replace(s, ".", "/")
    s = Replace(s, "\", "/")
    s = Replace(s, " ", "/")
    Do While InStr(s, "//") > 0
        s = Replace(s, "//", "/")
    Loop
    If InStr(s, "/") > 0 Then
        parts = Split(s, "/")
        If UBound(parts) <> 2 Then Exit Function
        If Len(parts(0)) = 4 Then
            TryBirthDate = ValidYMD(parts(0), parts(1), parts(2), d)
        ElseIf Len(parts(2)) = 4 Then
            TryBirthDate = ValidYMD(parts(2), parts(1), parts(0), d)
        End If
    ElseIf Len(s) = 8 And DigitsOnly(s) Then
        If ValidYMD(Left$(s, 4), Mid$(s, 5, 2), Right$(s, 2), d) Then
            TryBirthDate = True
        Else
            TryBirthDate = ValidYMD(Right$(s, 4), Mid$(s, 3, 2), Left$(s, 2), d)
        End If
    End If
InvalidDate:
End Function

Private Function PositiveGroup(ByVal s As String) As Boolean
    If Len(s) > 9 Then Exit Function
    If Not DigitsOnly(s) Then Exit Function
    PositiveGroup = (CLng(s) > 0)
End Function

Private Sub AddValidationIssue(ByVal problems As Object, ByVal location As String, ByVal detail As String)
    If problems.Exists(location) Then
        If InStr(1, " " & problems(location) & " ", " " & detail & " ", vbBinaryCompare) = 0 Then
            problems(location) = problems(location) & " " & detail
        End If
    Else
        problems.Add location, detail
    End If
End Sub

Private Sub RaiseValidationIssues(ByVal problems As Object, ByVal legend As String)
    Const MAX_ROWS As Long = 20
    Const MAX_LINE As Long = 40
    Dim key As Variant, shown As Long, line As String, msg As String
    If problems.Count = 0 Then Exit Sub
    msg = "Invalid rows: " & problems.Count & " (up to 20 shown)." & vbLf & legend
    For Each key In problems.Keys
        If shown >= MAX_ROWS Then Exit For
        line = CStr(key) & ": " & problems(key)
        If Len(line) > MAX_LINE Then line = Left$(line, MAX_LINE - 3) & "..."
        msg = msg & vbLf & line
        shown = shown + 1
    Next key
    msg = msg & vbLf & "No data written."
    Err.Raise vbObjectError + 58, , msg
End Sub

Private Function Txt(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = 0 To UBound(parts)
        s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    Txt = s
End Function

Private Function ArtsText(ByVal key As String) As String
    Select Case key
    Case "arts": ArtsText = Txt("622 62F 627 628 20 648 641 644 633 641 629")
    Case "lang": ArtsText = Txt("644 63A 627 62A 20 623 62C 646 628 64A 629")
    Case "f_arts": ArtsText = Txt("622 62F 627 628")
    Case "f_philo": ArtsText = Txt("641 644 633 641 629")
    Case "f_lang": ArtsText = Txt("644 63A 627 62A")
    Case "arts_prefix": ArtsText = Txt("623 648 644 649 20 62B 627 646 648 64A 20 62C 630 639 20 645 634 62A 631 643 20 622 62F 627 628")
    Case "science_report": ArtsText = Txt("62D 648 635 644 629 20 2D 20 639 644 648 645")
    Case "arts_report": ArtsText = Txt("62D 648 635 644 629 20 2D 20 622 62F 627 628")
    Case "legacy_report": ArtsText = Txt("627 644 62D 648 635 644 629")
    Case "arts_title": ArtsText = Txt("627 644 62A 648 62C 64A 647 20 627 644 62A 62F 631 64A 62C 64A 20 644 644 641 635 644 20 627 644 62B 627 646 64A 20 2D 20 623 648 644 649 20 62B 627 646 648 64A 20 62C 630 639 20 645 634 62A 631 643 20 622 62F 627 628")
    Case "arts_report_title": ArtsText = Txt("62D 635 64A 644 629 20 627 644 62A 648 62C 64A 647 20 627 644 62A 62F 631 64A 62C 64A 20 644 644 62C 630 639 20 627 644 645 634 62A 631 643 20 622 62F 627 628 20 62D 633 628 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A")
    Case "number": ArtsText = Txt("627 644 631 642 645")
    Case "name": ArtsText = Txt("627 644 644 642 628 20 648 627 644 627 633 645")
    Case "birth": ArtsText = Txt("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F")
    Case "sex": ArtsText = Txt("627 644 62C 646 633")
    Case "repeat": ArtsText = Txt("627 644 625 639 627 62F 629")
    Case "arabic": ArtsText = Txt("627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627")
    Case "history": ArtsText = Txt("627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627")
    Case "french": ArtsText = Txt("627 644 644 63A 629 20 627 644 641 631 646 633 64A 629")
    Case "english": ArtsText = Txt("627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629")
    Case "average": ArtsText = Txt("645 639 62F 644 20 627 644 641 635 644 20 32")
    Case "wish1": ArtsText = Txt("627 644 631 63A 628 629 20 30 31")
    Case "wish2": ArtsText = Txt("627 644 631 63A 628 629 20 30 32")
    Case "p_lang": ArtsText = Txt("645 644 645 62D 20 644 63A 627 62A 20 623 62C 646 628 64A 629")
    Case "p_arts": ArtsText = Txt("645 644 645 62D 20 622 62F 627 628 20 648 641 644 633 641 629")
    Case "decision": ArtsText = Txt("627 642 62A 631 627 62D 20 645 633 62A 634 627 631 20 627 644 62A 648 62C 64A 647")
    Case "invalidated": ArtsText = Txt("62A 63A 64A 651 631 62A 20 628 64A 627 646 627 62A 20 647 630 627 20 627 644 645 633 627 631 2E 20 64A 62C 628 20 625 639 627 62F 629 20 62A 634 63A 64A 644 20 627 644 62A 648 62C 64A 647 20 644 62A 62D 62F 64A 62B 20 627 644 642 631 627 631 627 62A 20 648 627 644 62D 648 635 644 629 2E")
    Case "failure_unsaved": ArtsText = Txt("642 62F 20 62A 643 648 646 20 628 64A 627 646 627 62A 20 627 644 645 644 641 20 627 644 645 641 62A 648 62D 20 642 62F 20 62A 63A 64A 651 631 62A 20 62C 632 626 64A 64B 627 2E 20 644 627 20 62A 62D 641 638 20 627 644 645 644 641 60C 20 648 623 63A 644 642 647 20 62F 648 646 20 62D 641 638 20 62B 645 20 623 639 62F 20 641 62A 62D 647 2E")
    Case "header_expected": ArtsText = Txt("627 644 639 646 648 627 646 20 627 644 645 62A 648 642 639 20 64A 62A 636 645 646 3A 20")
    Case "h_number": ArtsText = Txt("631 642 645")
    Case "h_name": ArtsText = Txt("627 633 645")
    Case "h_birth": ArtsText = Txt("645 64A 644 627 62F")
    Case "h_sex": ArtsText = Txt("62C 646 633")
    Case "h_repeat": ArtsText = Txt("625 639 627 62F 629")
    Case "h_arabic": ArtsText = Txt("639 631 628 64A")
    Case "h_math": ArtsText = Txt("631 64A 627 636 64A 627 62A")
    Case "h_physics": ArtsText = Txt("641 64A 632 64A 627")
    Case "h_natural": ArtsText = Txt("637 628 64A 639")
    Case "h_history": ArtsText = Txt("62A 627 631 64A 62E")
    Case "h_french": ArtsText = Txt("641 631 646 633 64A")
    Case "h_english": ArtsText = Txt("627 646 62C 644 64A 632")
    Case "h_info": ArtsText = Txt("645 639 644 648 645 627 62A")
    Case "h_info_alt": ArtsText = Txt("625 639 644 627 645")
    Case "h_technology": ArtsText = Txt("62A 643 646 648 644 648 62C")
    Case "h_average": ArtsText = Txt("645 639 62F 644 20 627 644 641 635 644")
    Case "h_avg_word": ArtsText = Txt("645 639 62F 644")
    Case "h_term_word": ArtsText = Txt("641 635 644")
    Case "general_report": ArtsText = Txt("62D 648 635 644 629 20 639 627 645 629")
    Case "general_title": ArtsText = Txt("62D 635 64A 644 629 20 627 644 62A 648 62C 64A 647 20 627 644 646 647 627 626 64A")
    Case "general_subtitle": ArtsText = Txt("627 642 62A 631 627 62D 627 62A 20 627 644 62A 648 62C 64A 647 20 644 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646")
    Case "track_science": ArtsText = Txt("62C 20 645 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627")
    Case "track_arts": ArtsText = Txt("62C 20 645 20 622 62F 627 628")
    Case "male": ArtsText = Txt("630 643 631")
    Case "female": ArtsText = Txt("623 646 62B 649")
    End Select
End Function

Private Function GuidanceTable(ByVal ws As Worksheet, ByVal width As Long) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo.Range.Column = 1 And lo.HeaderRowRange.Row = 2 Then
            If lo.ListColumns.Count < width Then Err.Raise vbObjectError + 80, , ws.Name & ": too few table columns."
            Set GuidanceTable = lo
            Exit Function
        End If
    Next lo
    Err.Raise vbObjectError + 80, , ws.Name & ": a table starting at A2 is required."
End Function

Private Function SummarySheet(ByVal wb As Workbook, ByVal isArts As Boolean, ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = FindWorksheet(wb, ArtsText("general_report"))
    If ws Is Nothing And create Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = ArtsText("general_report")
    End If
    Set SummarySheet = ws
End Function

Private Sub CheckTrackWrite(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim report As Worksheet
    ws.Unprotect Password:=""
    If ws.ProtectContents Then Err.Raise vbObjectError + 81, , "Unprotect " & ws.Name & " before running."
    Set report = SummarySheet(wb, False, False)
    If Not report Is Nothing Then
        report.Unprotect Password:=""
        If report.ProtectContents Then Err.Raise vbObjectError + 81, , "Unprotect " & report.Name & " before running."
    End If
End Sub

Private Sub LockAndProtect(ByVal ws As Worksheet, ByVal body As Range)
    Dim firstWish As Long
    If ws.Name = "Data01" Then firstWish = 11 Else firstWish = 15
    ws.Cells.Locked = True
    If Not body Is Nothing Then body.Columns(firstWish).Resize(, 2).Locked = False
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
End Sub

Private Sub InvalidateOrientation(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim report As Worksheet, lastRow As Long, decisionCol As Long
    If ws.Name = "Data01" Then decisionCol = 15 Else decisionCol = 21
    Set report = SummarySheet(wb, False, False)
    If Not report Is Nothing Then
        report.Unprotect Password:=""
        If report.ProtectContents Then Err.Raise vbObjectError + 81, , "Unprotect " & report.Name & " before running."
    End If
    lastRow = ws.Cells(ws.Rows.Count, decisionCol).End(xlUp).Row
    If lastRow >= 3 Then ws.Range(ws.Cells(3, decisionCol), ws.Cells(lastRow, decisionCol)).ClearContents
    If Not report Is Nothing Then report.Calculate
End Sub

Private Function ColumnLetter(ByVal n As Long) As String
    Do While n > 0
        ColumnLetter = Chr$(65 + (n - 1) Mod 26) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

Private Function TransferBatchMessage(ByRef batch As TransferBatch) As String
    Dim nm As String
    If batch.IsArts Then nm = "Data01" Else nm = "Data"
    If batch.SheetCount = 0 Then
        TransferBatchMessage = nm & ": no matching source track; skipped, existing data kept."
    Else
        TransferBatchMessage = nm & ": " & batch.Rows & " student(s) transferred; selected pairs averaged."
        If batch.ErrorCells > 0 Then TransferBatchMessage = TransferBatchMessage & " Ignored error cells: " & batch.ErrorCells & "."
        If batch.NonNumericCells > 0 Then TransferBatchMessage = TransferBatchMessage & " Ignored non-numeric cells: " & batch.NonNumericCells & "."
    End If
End Function

Private Sub ValidateSourceHeaders(ByVal ws As Worksheet, ByVal isArts As Boolean, ByVal map As Variant, ByVal problems As Object)
    Dim keys As Variant, i As Long, c As Long, header As Variant, key As String, k As Long, fieldWidth As Long
    If isArts Then
        keys = Array("h_number", "h_name", "h_birth", "h_sex", "h_repeat", "h_arabic", "h_history", "h_french", "h_english", "h_average")
    Else
        keys = Array("h_number", "h_name", "h_birth", "h_sex", "h_repeat", "h_arabic", "h_math", "h_physics", "h_natural", "h_history", "h_info", "h_technology", "h_average")
    End If
    ws.Range(ws.Cells(2, 1), ws.Cells(2, CLng(map(UBound(map))) + GROUP_SIZE - 1)).Calculate
    For i = LBound(map) To UBound(map)
        c = CLng(map(i))
        key = CStr(keys(i))
        fieldWidth = 1
        If i >= 5 Then fieldWidth = GROUP_SIZE
        For k = 0 To fieldWidth - 1
            header = ws.Cells(2, c + k).MergeArea.Cells(1, 1).Value2
            If Not SourceHeaderMatches(header, key) Then
                AddValidationIssue problems, ws.Name & "!" & ColumnLetter(c + k) & "2", ArtsText("header_expected") & ArtsText(key)
            End If
        Next k
    Next i
End Sub

Private Function SourceHeaderMatches(ByVal value As Variant, ByVal key As String) As Boolean
    Dim h As String
    If IsError(value) Or IsNull(value) Then Exit Function
    h = Norm(value)
    If key = "h_average" Then
        SourceHeaderMatches = InStr(h, Norm(ArtsText("h_avg_word"))) > 0 And InStr(h, Norm(ArtsText("h_term_word"))) > 0
    ElseIf key = "h_info" Then
        SourceHeaderMatches = InStr(h, Norm(ArtsText(key))) > 0 Or InStr(h, Norm(ArtsText("h_info_alt"))) > 0
    Else
        SourceHeaderMatches = InStr(h, Norm(ArtsText(key))) > 0
    End If
End Function

Private Function FindWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FindWorksheet = wb.Worksheets(sheetName)
    Err.Clear
    On Error GoTo 0
End Function

Private Function RequireWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = FindWorksheet(wb, sheetName)
    If ws Is Nothing Then Err.Raise vbObjectError + 83, , "Required sheet '" & sheetName & "' not found in " & wb.Name & ". Run data transfer first."
    Set RequireWorksheet = ws
End Function

Private Function AveragePair(ByRef sourceData As Variant, ByVal rowNum As Long, _
                             ByVal firstCol As Long, ByRef errorCells As Long, _
                             ByRef nonNumericCells As Long) As Variant
    Dim k As Long, c As Long, countVal As Long
    Dim v As Variant, sumVal As Double
    For k = 0 To GROUP_SIZE - 1
        c = firstCol + k
        If c <= UBound(sourceData, 2) Then
            v = sourceData(rowNum, c)
            If Not IsError(v) Then
                If Not IsEmpty(v) Then
                    If IsNumeric(v) Then
                        sumVal = sumVal + CDbl(v)
                        countVal = countVal + 1
                    ElseIf Len(Trim$(CStr(v))) > 0 Then
                        nonNumericCells = nonNumericCells + 1
                    End If
                End If
            Else
                errorCells = errorCells + 1
            End If
        End If
    Next k
    If countVal > 0 Then AveragePair = sumVal / countVal
End Function

' ==============================================================================
' EMBEDDED COLLECTION ENGINE
' ==============================================================================
Private Function FileExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FileExistsPath = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExistsPath = False: Err.Clear
End Function

Private Function SortedNames(ByVal col As Collection) As Variant
    Dim arr() As String, i As Long, j As Long, tmp As String
    If col.Count = 0 Then SortedNames = Array(): Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count: arr(i) = CStr(col(i)): Next i
    For i = 2 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 1
            If StrComp(arr(j), tmp, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    SortedNames = arr
End Function

Private Function FontExists(ByVal FontName As String) As Boolean
    Const HKCU As Long = &H80000001
    Const HKLM As Long = &H80000002
    Dim reg As Object, hv As Variant
    Dim names As Variant, types As Variant
    Dim i As Long, p As Long, nm As String, enumOK As Boolean

    On Error GoTo Fallback
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    For Each hv In Array(HKLM, HKCU)
        names = Empty
        If reg.EnumValues(CLng(hv), "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", names, types) = 0 Then
            enumOK = True
            If IsArray(names) Then
                For i = LBound(names) To UBound(names)
                    nm = CStr(names(i))
                    p = InStr(nm, " (")
                    If p > 0 Then nm = Left(nm, p - 1)
                    If StrComp(nm, FontName, vbTextCompare) = 0 Then FontExists = True: Exit Function
                    If StrComp(Left(FontName, Len(nm) + 1), nm & " ", vbTextCompare) = 0 Then FontExists = True: Exit Function
                Next i
            End If
        End If
    Next hv

    FontExists = Not enumOK
    Exit Function
Fallback:
    FontExists = True
End Function

Private Sub SetupFonts()
    If FontExists(FONT_ARABIC) Then g_ActiveArabicFont = FONT_ARABIC Else g_ActiveArabicFont = FONT_FALLBACK
    If FontExists(FONT_NUMERIC) Then g_ActiveNumericFont = FONT_NUMERIC Else g_ActiveNumericFont = FONT_FALLBACK
End Sub

Private Function SrcProcessFile(ByRef wbResult As Workbook, _
                                ByVal SourcePath As String, _
                                ByVal fileName As String, _
                                ByVal SheetNumber As Long, _
                                ByRef ReasonOut As String) As Boolean
    Dim wbSource As Workbook, wsSource As Worksheet, wsResult As Worksheet
    Dim LastRow As Long, LastCol As Long, DataArr As Variant
    Dim errNum As Long, errDesc As String, lastDataRow As Long, wrapEndRow As Long
    Dim idLastCol As Long, lastFind As Range
    Dim tbl As ListObject, tblRange As Range, seenHeaders As Object
    Dim h As Long, headerText As String, genCounter As Long

    SrcProcessFile = False
    ReasonOut = ""
    On Error GoTo EH

    On Error Resume Next
    Set wbSource = Workbooks.Open(fileName:=SourcePath & fileName, UpdateLinks:=False, _
                                  ReadOnly:=True, AddToMRU:=False, Password:="")
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wbSource Is Nothing Or errNum <> 0 Then
        ReasonOut = "OPEN_FAILED | " & errDesc: Exit Function
    End If

    If wbSource.Sheets.Count = 0 Then
        ReasonOut = "EMPTY_WORKBOOK (no sheets)": wbSource.Close SaveChanges:=False: Exit Function
    End If

    Set wsSource = wbSource.Sheets(1)

    On Error Resume Next
    LastCol = wsSource.Cells(DATA_START_ROW, wsSource.Columns.Count).End(xlToLeft).Column
    Set lastFind = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                                       SearchDirection:=xlPrevious)
    If lastFind Is Nothing Then LastRow = 0 Else LastRow = lastFind.Row - TRAILING_ROWS_SKIP
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Then
        ReasonOut = "DETECTION_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    If LastRow < DATA_START_ROW Or LastCol < 1 Then
        ReasonOut = "EMPTY_WORKBOOK (no data at row " & DATA_START_ROW & ")"
        wbSource.Close SaveChanges:=False: Exit Function
    End If

    On Error Resume Next
    Set wsResult = wbResult.Sheets.Add(After:=wbResult.Sheets(wbResult.Sheets.Count))
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wsResult Is Nothing Then
        ReasonOut = "ADD_SHEET_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    On Error Resume Next
    wsResult.Name = Format(SheetNumber, "00")
    Err.Clear
    On Error GoTo EH

    wsResult.DisplayRightToLeft = True

    On Error Resume Next
    DataArr = wsSource.Range(wsSource.Cells(DATA_START_ROW, 1), wsSource.Cells(LastRow, LastCol)).Value
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Or Not IsArray(DataArr) Then
        ReasonOut = "DATA_READ_FAILED | " & errDesc
        wbSource.Close SaveChanges:=False
        wsResult.Delete
        Exit Function
    End If

    lastDataRow = 1 + UBound(DataArr, 1)

    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        If LastCol >= 4 Then .Columns("D").NumberFormat = "@"
        If LastCol >= 5 Then .Columns("E").NumberFormat = "@"
        If LastCol >= 6 Then .Range(.Columns("F"), .Columns(LastCol)).NumberFormat = "0.00"
    End With

    wsResult.Range("A2").Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    genCounter = 0

    For h = 1 To LastCol
        headerText = Trim(CStr(wsResult.Cells(2, h).Value))
        If headerText = "" Or seenHeaders.Exists(headerText) Then
            genCounter = genCounter + 1
            Do While seenHeaders.Exists("Column" & genCounter): genCounter = genCounter + 1: Loop
            wsResult.Cells(2, h).Value = "Column" & genCounter
            seenHeaders.Add "Column" & genCounter, True
        Else
            seenHeaders.Add headerText, True
        End If
    Next h

    Set tblRange = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(lastDataRow, LastCol))
    Set tbl = Nothing

    On Error Resume Next
    Set tbl = wsResult.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number = 0 Then
        tbl.Name = "Data" & Format(SheetNumber, "00")
        tbl.TableStyle = ""
        tbl.ShowTableStyleColumnStripes = False
        tbl.ShowTableStyleRowStripes = False
        tbl.ShowTableStyleFirstColumn = False
        tbl.ShowTableStyleLastColumn = False
        tbl.ShowAutoFilter = True
    End If
    Err.Clear
    On Error GoTo EH

    With wsResult
        .Columns("B").Font.Name = g_ActiveArabicFont: .Columns("B").Font.Size = FONT_SIZE
        .Columns("C").Font.Name = g_ActiveNumericFont: .Columns("C").Font.Size = FONT_SIZE
        If LastCol >= 4 Then .Columns("D").Font.Name = g_ActiveArabicFont: .Columns("D").Font.Size = FONT_SIZE
        If LastCol >= 5 Then .Columns("E").Font.Name = g_ActiveArabicFont: .Columns("E").Font.Size = FONT_SIZE
        If LastCol >= 6 Then
            With .Range(.Columns("F"), .Columns(LastCol))
                .Font.Name = g_ActiveNumericFont
                .Font.Size = FONT_SIZE
            End With
        End If

        wrapEndRow = lastDataRow
        If wrapEndRow < 2 Then wrapEndRow = 2
        .Range(.Cells(2, 1), .Cells(wrapEndRow, LastCol)).WrapText = False

        With .Range(.Cells(1, 1), .Cells(lastDataRow, LastCol))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        idLastCol = IDENTITY_LAST_COL
        If LastCol < idLastCol Then idLastCol = LastCol
        .Range(.Columns(1), .Columns(idLastCol)).EntireColumn.AutoFit

        If LastCol >= 6 Then .Range(.Columns(IDENTITY_LAST_COL + 1), .Columns(LastCol)).ColumnWidth = COL_WIDTH

        If LastCol >= 4 Then
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Orientation = xlUpward
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Font.Bold = True
        End If
    End With

    PlaceHeaderTitle wsSource, wsResult

    On Error Resume Next
    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing
    Err.Clear
    On Error GoTo 0

    SrcProcessFile = True
    Exit Function

EH:
    ReasonOut = "UNEXPECTED | " & Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If Not wsResult Is Nothing Then wsResult.Delete
    Err.Clear
    SrcProcessFile = False
End Function

Private Sub SrcCollectAll(ByRef wbResult As Workbook, _
                          ByVal SourcePath As String, _
                          ByRef CopiedCount As Long, _
                          ByRef SkippedCount As Long, _
                          ByRef SkippedList As String)
    Dim fileName As String, SheetNumber As Long, Reason As String
    Dim names As Collection, sorted As Variant, i As Long
    Dim f As String, ext As String, baseName As String, dotPos As Long

    CopiedCount = 0: SkippedCount = 0: SheetNumber = 1: SkippedList = ""

    Set names = New Collection
    fileName = Dir(SourcePath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            dotPos = InStrRev(fileName, ".")
            ext = LCase(Mid(fileName, dotPos + 1))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Then names.Add fileName
        End If
        fileName = Dir
    Loop

    If names.Count = 0 Then
        SkippedList = "  - " & SRC_FOLDER & " folder is empty or contains no Excel files:" & vbCrLf & _
                      "    " & SourcePath & vbCrLf
        Exit Sub
    End If

    sorted = SortedNames(names)
    SetupFonts

    For i = LBound(sorted) To UBound(sorted)
        f = CStr(sorted(i))
        dotPos = InStrRev(f, ".")
        ext = LCase(Mid(f, dotPos + 1))
        baseName = Left(f, dotPos - 1)

        If StrComp(SourcePath & f, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        ElseIf IsOpenInExcel(SourcePath & f) Then
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : ALREADY OPEN in Excel (close it and run again)" & vbCrLf
        ElseIf ext <> "xlsx" And FileExistsPath(SourcePath & baseName & ".xlsx") Then
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : DUPLICATE (.xlsx version exists)" & vbCrLf
        ElseIf SrcProcessFile(wbResult, SourcePath, f, SheetNumber, Reason) Then
            CopiedCount = CopiedCount + 1: SheetNumber = SheetNumber + 1
        Else
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : " & Reason & vbCrLf
        End If
    Next i
End Sub

Private Function IsOpenInExcel(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsOpenInExcel = True
            Exit Function
        End If
    Next wb
End Function

Private Function SrcBuildWorkbook(ByVal srcFolder As String, _
                                  ByRef wbOut As Workbook, _
                                  ByRef copied As Long, _
                                  ByRef skipped As Long, _
                                  ByRef skippedList As String) As Boolean
    Dim wb As Workbook, oldSec As Long, oldAlerts As Boolean

    SrcBuildWorkbook = False
    Set wbOut = Nothing
    copied = 0: skipped = 0: skippedList = ""

    If LCase(Left(srcFolder, 4)) = "http" Then
        skippedList = "  - The workbook is on OneDrive/SharePoint (URL path). Move it to a local folder first." & vbCrLf
        Exit Function
    End If
    If Right(srcFolder, 1) <> "\" Then srcFolder = srcFolder & "\"
    If Dir(srcFolder, vbDirectory) = "" Then
        skippedList = "  - " & SRC_FOLDER & " folder not found: " & srcFolder & vbCrLf
        Exit Function
    End If

    oldSec = Application.AutomationSecurity
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    On Error GoTo EH

    Set wb = Workbooks.Add
    Do While wb.Sheets.Count > 1: wb.Sheets(wb.Sheets.Count).Delete: Loop

    SrcCollectAll wb, srcFolder, copied, skipped, skippedList

    If copied > 0 Then
        If wb.Sheets.Count > 1 Then
            If Left(wb.Sheets(1).Name, 5) = "Sheet" And wb.Sheets(1).UsedRange.Cells.Count = 1 Then wb.Sheets(1).Delete
        End If
        Set wbOut = wb
        SrcBuildWorkbook = True
    Else
        wb.Close SaveChanges:=False
    End If

Done:
    Application.AutomationSecurity = oldSec
    Application.DisplayAlerts = oldAlerts
    Exit Function

EH:
    skippedList = skippedList & "  - Build failed: " & Err.Number & " : " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wbOut = Nothing
    SrcBuildWorkbook = False
    Resume Done
End Function

Private Sub PlaceHeaderTitle(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet)
    Dim rawText As String
    Dim partTitle As String
    Dim partGroup As String

    rawText = CleanText(CStr(wsSource.Range("A5").Value))
    SplitAtSchoolYear rawText, partTitle, partGroup

    WriteMergedCell wsTarget.Range("B1:F1"), partTitle
    WriteMergedCell wsTarget.Range("H1:Q1"), partGroup

    wsTarget.Rows(1).RowHeight = HEADER_ROW_HEIGHT
End Sub

Private Function CleanText(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(160), " ")
    s = Application.WorksheetFunction.Clean(s)
    CleanText = Application.WorksheetFunction.Trim(s)
End Function

Private Sub SplitAtSchoolYear(ByVal txt As String, _
                              ByRef part1 As String, _
                              ByRef part2 As String)
    Dim rx As Object
    Dim m As Object
    Dim splitPos As Long
    Dim digit As String
    Dim seps As String

    part1 = txt
    part2 = ""
    If Len(txt) = 0 Then Exit Sub

    digit = "[0-9\u0660-\u0669]"
    seps = "[/\\\-\u2013\u2014]"

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True

    rx.Pattern = digit & "{4}\s*" & seps & "\s*" & digit & "{4}"

    If Not rx.Test(txt) Then
        rx.Pattern = digit & "{4}"
        If Not rx.Test(txt) Then Exit Sub
    End If

    Set m = rx.Execute(txt)
    splitPos = m(0).FirstIndex + m(0).Length

    part1 = Trim(Left(txt, splitPos))

    If splitPos < Len(txt) Then
        part2 = Trim(Mid(txt, splitPos + 1))

        Do While Len(part2) > 0
            If InStr(1, "-:/" & ChrW(&H2013) & ChrW(&H2014), Left(part2, 1)) = 0 Then Exit Do
            part2 = Trim(Mid(part2, 2))
        Loop
    End If
End Sub

Private Sub WriteMergedCell(ByVal rng As Range, ByVal txt As String)
    With rng
        .UnMerge
        .ClearContents
        .Merge
        .Value = txt
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
        .Font.Name = GetAvailableFont(PREFERRED_FONT, FALLBACK_FONT)
        .Font.Size = HEADER_FONT_SIZE
        .Font.Bold = True
    End With
End Sub

Private Function GetAvailableFont(ByVal wanted As String, _
                                  ByVal fallback As String) As String
    If IsFontInstalled(wanted) Then
        GetAvailableFont = wanted
    Else
        GetAvailableFont = fallback
    End If
End Function

Private Function IsFontInstalled(ByVal fontName As String) As Boolean
    Dim folders(1 To 2) As String
    Dim i As Long

    folders(1) = Environ("WINDIR") & "\Fonts\"
    folders(2) = Environ("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\"

    On Error Resume Next
    For i = 1 To 2
        If Len(Dir(folders(i) & fontName & "*.*")) > 0 Then
            IsFontInstalled = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function
