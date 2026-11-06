Attribute VB_Name = "modTransfer02"
Option Explicit

Private m_runSucceeded As Boolean
Private m_lastReport As String
Private m_silent As Boolean
Private m_guidanceWb As Workbook     ' in-memory guidance workbook shared with the next modules
Private m_guidanceSaved As String    ' path of the saved copy ("" while never saved)

' ============================================================
' Transfer 02: \Source02\ -> Term02Guidance workbook (Data science; Data01 arts)
' The guidance workbook is built from scratch (design copied from the model file) and stays
' UNSAVED in memory: modStreamPreferences02 and modOrientation02 continue on it, and
' SaveGuidanceCopy stores it in the output folder next to the tool workbook (created when missing).
' ============================================================

Private Const SRC_FOLDER As String = "Source02"
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
    Dim wbSrc As Workbook, wbDst As Workbook, srcOpened As Boolean
    Dim science As TransferBatch, arts As TransferBatch, problems As Object
    Dim failed As Boolean, errText As String
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
    ReleaseGuidanceWorkbook
    BuildGuidanceWorkbook wbDst
    Set problems = CreateObject("Scripting.Dictionary")
    PrepareTransfer wbSrc, wbDst, False, science, problems
    PrepareTransfer wbSrc, wbDst, True, arts, problems
    RaiseValidationIssues problems, "Header or data errors: sheet, cell/row and expected column are listed."
    If science.SheetCount + arts.SheetCount = 0 Then Err.Raise vbObjectError + 3, , "No matching science or arts sheet found."

    If science.SheetCount > 0 And science.Rows > 0 Then WriteTransfer wbDst, science Else FinishEmptyTrack wbDst, False
    If arts.SheetCount > 0 And arts.Rows > 0 Then WriteTransfer wbDst, arts Else FinishEmptyTrack wbDst, True
    Set m_guidanceWb = wbDst
    m_guidanceSaved = ""
    m_runSucceeded = True
    m_lastReport = TransferBatchMessage(science) & vbLf & TransferBatchMessage(arts) & science.Skipped & arts.Skipped
    ' Run on its own (not part of the orientation sequence): keep a copy so the work is not lost
    If Not wasSilent Then m_lastReport = m_lastReport & vbLf & "Saved to: " & SaveGuidanceCopy()

Done:
    On Error Resume Next
    If srcOpened Then wbSrc.Close SaveChanges:=False
    If failed Then
        If Not wbDst Is Nothing Then wbDst.Close SaveChanges:=False
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

Private Sub WriteTransfer(ByVal wb As Workbook, ByRef batch As TransferBatch)
    Dim ws As Worksheet, lo As ListObject, body As Range, width As Long
    If batch.SheetCount = 0 Or batch.Rows = 0 Then Exit Sub
    If batch.IsArts Then
        Set ws = wb.Worksheets("Data01"): width = 15
    Else
        Set ws = wb.Worksheets("Data"): width = 21
    End If
    Set lo = GuidanceTable(ws, width)
    lo.Resize ws.Range(ws.Cells(2, 1), ws.Cells(2 + batch.Rows, lo.ListColumns.Count))
    Set body = lo.DataBodyRange
    body.Resize(, width).Value2 = batch.Data
    FormatGuidanceBody ws, batch.IsArts, 2 + batch.Rows
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
' IN-MEMORY GUIDANCE WORKBOOK
' Built from scratch by Run_Transfer02 (sheets Data and Data01, design copied from the
' model file). It stays UNSAVED in memory while modStreamPreferences02 and modOrientation02
' work on it; SaveGuidanceCopy stores it in the output folder next to the tool workbook.
' ==============================================================================

' The workbook built by the last successful Run_Transfer02 (Nothing when none is open)
Public Function GuidanceWorkbook() As Workbook
    Dim nm As String
    If m_guidanceWb Is Nothing Then Exit Function
    On Error Resume Next
    nm = m_guidanceWb.Name
    If Err.Number <> 0 Then
        Err.Clear
        Set m_guidanceWb = Nothing
        m_guidanceSaved = ""
    End If
    On Error GoTo 0
    Set GuidanceWorkbook = m_guidanceWb
End Function

' Full path of the saved copy ("" while the workbook has never been saved)
Public Function GuidanceSavedPath() As String
    GuidanceSavedPath = m_guidanceSaved
End Function

' Saves the in-memory workbook as Term02Guidance[_NN].xlsx inside the output folder
' (folder created when missing; an existing file is never replaced). A second call
' in the same run simply saves the same file again.
Public Function SaveGuidanceCopy() As String
    Dim wb As Workbook, outDir As String, p As String
    Set wb = GuidanceWorkbook()
    If wb Is Nothing Then Err.Raise vbObjectError + 20, , "No guidance workbook is open. Run the data transfer first."
    If Len(m_guidanceSaved) = 0 Then
        If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 1, , "Save the tool workbook before running this procedure."
        outDir = ThisWorkbook.Path & Application.PathSeparator & OutputFolderName()
        If Not EnsureFolder(outDir) Then Err.Raise vbObjectError + 21, , "The output folder could not be created: " & outDir
        p = NextGuidancePath(outDir)
        wb.SaveAs Filename:=p, FileFormat:=xlOpenXMLWorkbook
        m_guidanceSaved = p
    Else
        wb.Save
    End If
    SaveGuidanceCopy = m_guidanceSaved
End Function

Private Sub ReleaseGuidanceWorkbook()
    Dim wb As Workbook
    Set wb = GuidanceWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    Err.Clear
    Set m_guidanceWb = Nothing
    m_guidanceSaved = ""
    On Error GoTo 0
End Sub

Private Function NextGuidancePath(ByVal folderPath As String) As String
    Dim cand As String, seq As Long
    If Right$(folderPath, 1) <> Application.PathSeparator Then folderPath = folderPath & Application.PathSeparator
    cand = "Term02Guidance.xlsx"
    Do While FileExistsPath(folderPath & cand) Or WorkbookNameOpen(cand)
        seq = seq + 1
        cand = "Term02Guidance_" & Format$(seq, "00") & ".xlsx"
    Loop
    NextGuidancePath = folderPath & cand
End Function

Private Function WorkbookNameOpen(ByVal fileName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            WorkbookNameOpen = True
            Exit Function
        End If
    Next wb
End Function

' ==============================================================================
' BUILDER (layout, fonts, colors, widths and print setup copied from the model file)
' ==============================================================================
Private Sub BuildGuidanceWorkbook(ByRef wb As Workbook)
    Dim wsS As Worksheet, wsA As Worksheet
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error Resume Next
    With wb.Styles("Normal").Font
        .Name = "Calibri Light"
        .Size = 14
    End With
    Err.Clear
    On Error GoTo 0
    Set wsS = wb.Worksheets(1)
    wsS.Name = "Data"
    Set wsA = wb.Worksheets.Add(After:=wsS)
    wsA.Name = "Data01"
    BuildGuidanceSheet wsS, False
    BuildGuidanceSheet wsA, True
    On Error Resume Next
    wsS.Activate
    wsS.Range("A1").Select
    On Error GoTo 0
End Sub

Private Function GFont(ByVal wanted As String, ByVal fallback As String) As String
    Static cache As Object          ' the font check is slow: ask once per font
    If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
    If Not cache.Exists(wanted) Then
        If FontExists(wanted) Then cache(wanted) = wanted Else cache(wanted) = fallback
    End If
    GFont = CStr(cache(wanted))
End Function

Private Function GuidanceHeaders(ByVal isArts As Boolean) As Variant
    Dim h() As String
    If isArts Then
        ReDim h(0 To 14)
        h(0) = Txt("627 644 631 642 645")
        h(1) = Txt("627 644 644 642 628 20 648 20 627 644 627 633 645")
        h(2) = Txt("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F")
        h(3) = Txt("627 644 62C 646 633")
        h(4) = Txt("627 644 625 639 627 62F 629")
        h(5) = Txt("627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627")
        h(6) = Txt("627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627")
        h(7) = Txt("627 644 644 63A 629 20 627 644 641 631 646 633 64A 629")
        h(8) = Txt("627 644 644 63A 629 20 627 644 627 646 62C 644 64A 632 64A 629")
        h(9) = Txt("645 639 62F 644 20 627 644 641 635 644 20 32")
        h(10) = Txt("627 644 631 63A 628 629 20 30 31")
        h(11) = Txt("627 644 631 63A 628 629 20 30 32")
        h(12) = Txt("645 644 645 62D 20 644 63A 627 62A 20 623 62C 646 628 64A 629")
        h(13) = Txt("645 644 645 62D 20 622 62F 627 628 20 648 641 644 633 641 629")
        h(14) = Txt("627 642 62A 631 627 62D 20 645 633 62A 634 627 631 20 627 644 62A 648 62C 64A 647")
    Else
        ReDim h(0 To 20)
        h(0) = Txt("627 644 631 642 645")
        h(1) = Txt("627 644 644 642 628 20 648 627 644 627 633 645")
        h(2) = Txt("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F")
        h(3) = Txt("627 644 62C 646 633")
        h(4) = Txt("627 644 625 639 627 62F 629")
        h(5) = Txt("627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627")
        h(6) = Txt("627 644 631 64A 627 636 64A 627 62A")
        h(7) = Txt("627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629")
        h(8) = Txt("627 644 639 644 648 645 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629")
        h(9) = Txt("627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627")
        h(10) = Txt("627 644 645 639 644 648 645 627 62A 64A 629")
        h(11) = Txt("62A 643 646 648 644 648 62C 64A 627")
        h(12) = Txt("645 639 62F 644 20 627 644 641 635 644 20 32")
        h(13) = Txt("627 644 642 633 645")
        h(14) = Txt("627 644 631 63A 628 629 20 30 31")
        h(15) = Txt("627 644 631 63A 628 629 20 30 32")
        h(16) = Txt("645 644 645 62D 20 631 64A 627 636 64A 627 62A")
        h(17) = Txt("645 644 645 62D 20 62A 642 646 64A 20 631 64A 627 636 64A")
        h(18) = Txt("645 644 645 62D 20 639 644 648 645 20 62A 62C 631 64A 628 64A 629")
        h(19) = Txt("645 644 645 62D 20 62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F")
        h(20) = Txt("627 642 62A 631 627 62D 20 627 644 62A 648 62C 64A 647 20 627 644 645 633 628 642")
    End If
    GuidanceHeaders = h
End Function

Private Function GuidanceTitle(ByVal isArts As Boolean) As String
    If isArts Then
        GuidanceTitle = ArtsText("arts_title")
    Else
        GuidanceTitle = Txt("645 62D 636 631 20 627 644 62A 648 62C 64A 647 20 627 644 62A 62F 631 64A 62C 64A 20 62D 633 628 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A 20 644 644 633 646 629 20 627 644 623 648 644 649 20 62C 630 639 20 645 634 62A 631 643 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627")
    End If
End Function

' Formats a block: font, alignment, wrap, rotation, number format, lock, fill (fillRGB < 0 = none), thin borders
Private Sub FmtBlock(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c1 As Long, ByVal c2 As Long, _
                     ByVal fnt As String, ByVal isBold As Boolean, ByVal wrapIt As Boolean, ByVal orient As Long, _
                     ByVal nf As String, ByVal isLocked As Boolean, ByVal fillRGB As Long)
    With ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
        .Font.Name = fnt
        .Font.Size = 14
        .Font.Bold = isBold
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = wrapIt
        .Orientation = orient
        If Len(nf) > 0 Then .NumberFormat = nf
        .Locked = isLocked
        If fillRGB >= 0 Then
            .Interior.Color = fillRGB
        Else
            .Interior.Pattern = xlNone
        End If
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub BuildGuidanceSheet(ByVal ws As Worksheet, ByVal isArts As Boolean)
    Dim nc As Long, hdr As Variant, c As Long, lo As ListObject
    Dim arial As String, white As Long, f As String
    If isArts Then nc = 15 Else nc = 21
    white = RGB(255, 255, 255)
    arial = "Arial"
    hdr = GuidanceHeaders(isArts)
    ws.DisplayRightToLeft = True

    ' ---- table: header row 2 + one empty body row (resized by WriteTransfer) ----
    For c = 1 To nc
        ws.Cells(2, c).Value2 = hdr(c - 1)
    Next c
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(2, 1), ws.Cells(3, nc)), , xlYes)
    If isArts Then lo.Name = "Data01" Else lo.Name = "Data"
    lo.TableStyle = ""

    ' ---- title (row 1, merged over the table width) ----
    ws.Cells(1, 1).Value2 = GuidanceTitle(isArts)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nc))
        .Merge
        .Font.Name = GFont("bader_al merkaz", arial)
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ---- header formatting per column group ----
    If isArts Then
        FmtBlock ws, 2, 2, 1, 3, arial, True, False, 0, "", True, white
        FmtBlock ws, 2, 2, 4, 5, arial, True, False, 90, "", True, white
        FmtBlock ws, 2, 2, 6, 15, arial, True, True, 90, "", True, white
        ws.Rows(1).RowHeight = 36
        ws.Rows(2).RowHeight = 105.6
        ws.Columns(1).ColumnWidth = 5.1
        ws.Columns(2).ColumnWidth = 18.5
        ws.Columns(3).ColumnWidth = 10.9
        ws.Range("D:E").ColumnWidth = 4.6
        ws.Range("F:H").ColumnWidth = 5.8
        ws.Columns(9).ColumnWidth = 8.2
        ws.Columns(10).ColumnWidth = 7.7
        ws.Range("K:L").ColumnWidth = 10
        ws.Range("M:N").ColumnWidth = 9.3
        ws.Columns(15).ColumnWidth = 10
    Else
        f = GFont("Basha 8C", arial)
        FmtBlock ws, 2, 2, 1, 3, arial, True, False, 0, "", True, white
        FmtBlock ws, 2, 2, 4, 14, arial, True, True, 90, "", True, white
        FmtBlock ws, 2, 2, 15, 16, f, False, False, 90, "", True, white
        FmtBlock ws, 2, 2, 17, 21, f, False, True, 90, "", True, white
        ws.Rows(1).RowHeight = 26.4
        ws.Rows(2).RowHeight = 102
        ws.Columns(1).ColumnWidth = 5.1
        ws.Columns(2).ColumnWidth = 19.8
        ws.Columns(3).ColumnWidth = 11.7
        ws.Range("D:E").ColumnWidth = 6.6
        ws.Columns(6).ColumnWidth = 6
        ws.Range("G:H").ColumnWidth = 6.4
        ws.Range("I:N").ColumnWidth = 5.8
        ws.Range("O:P").ColumnWidth = 9.8
        ws.Range("Q:T").ColumnWidth = 6.6
        ws.Columns(21).ColumnWidth = 12.4
        BuildSideBlock ws, CStr(hdr(12))
    End If

    FormatGuidanceBody ws, isArts, 3
    GuidancePageSetup ws, isArts

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = isArts
    If isArts Then ActiveWindow.Zoom = 90
    On Error GoTo 0
End Sub

' Body formatting for rows 3..lastRow (called again after the table is resized)
Private Sub FormatGuidanceBody(ByVal ws As Worksheet, ByVal isArts As Boolean, ByVal lastRow As Long)
    Dim f As String, w As Long, r As Long
    w = RGB(255, 255, 255)
    If lastRow < 3 Then lastRow = 3
    If isArts Then
        f = GFont("Dubai", "Arial")
        FmtBlock ws, 3, lastRow, 1, 1, f, False, False, 0, "General", True, w
        FmtBlock ws, 3, lastRow, 2, 2, f, False, False, 0, "@", True, w
        FmtBlock ws, 3, lastRow, 3, 3, f, False, False, 0, "yyyy\-mm\-dd", True, w
        FmtBlock ws, 3, lastRow, 4, 5, f, False, False, 0, "@", True, w
        FmtBlock ws, 3, lastRow, 6, 7, f, False, False, 0, "0.00", True, w
        FmtBlock ws, 3, lastRow, 8, 8, f, False, False, 0, "0.00", True, RGB(180, 198, 231)
        FmtBlock ws, 3, lastRow, 9, 9, f, False, False, 0, "0.00", True, RGB(248, 203, 173)
        FmtBlock ws, 3, lastRow, 10, 10, f, False, False, 0, "0.00", True, RGB(198, 224, 180)
        FmtBlock ws, 3, lastRow, 11, 12, f, False, False, 0, "General", False, w
        FmtBlock ws, 3, lastRow, 13, 13, f, False, False, 0, "00.00", True, RGB(198, 224, 180)
        FmtBlock ws, 3, lastRow, 14, 14, f, False, False, 0, "00.00", True, RGB(219, 219, 219)
        FmtBlock ws, 3, lastRow, 15, 15, f, False, False, 0, "General", False, w
        ws.Rows(3 & ":" & lastRow).RowHeight = 19.95
        AddEmptyWishFormat ws.Range(ws.Cells(3, 11), ws.Cells(lastRow, 12))
        AddEmptyWishFormat ws.Range(ws.Cells(3, 15), ws.Cells(lastRow, 15))
    Else
        f = "Arial"
        FmtBlock ws, 3, lastRow, 1, 1, f, False, False, 0, "General", True, -1
        FmtBlock ws, 3, lastRow, 2, 2, f, False, False, 0, "@", True, -1
        FmtBlock ws, 3, lastRow, 3, 3, f, False, False, 0, "yyyy\-mm\-dd", True, -1
        FmtBlock ws, 3, lastRow, 4, 5, f, False, False, 0, "@", True, -1
        FmtBlock ws, 3, lastRow, 6, 13, f, False, False, 0, "0.00", True, -1
        FmtBlock ws, 3, lastRow, 14, 14, f, False, False, 0, "00", True, -1
        FmtBlock ws, 3, lastRow, 15, 16, GFont("Basha 8C", "Arial"), False, False, 0, "General", False, -1
        FmtBlock ws, 3, lastRow, 17, 20, f, False, False, 0, "00.00", True, -1
        FmtBlock ws, 3, lastRow, 21, 21, "Calibri", False, False, 0, "General", False, -1
        ws.Rows(3 & ":" & lastRow).RowHeight = 23.4
        AddEmptyWishFormat ws.Range(ws.Cells(3, 15), ws.Cells(lastRow, 16))
        AddEmptyWishFormat ws.Range(ws.Cells(3, 21), ws.Cells(lastRow, 21))
    End If
End Sub

' Cells that read "empty wish" are highlighted (bold red on yellow), as in the model file
Private Sub AddEmptyWishFormat(ByVal rng As Range)
    Dim fc As Object
    rng.FormatConditions.Delete
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & Txt("641 627 631 63A 629") & """")
    fc.Font.Bold = True
    fc.Font.Color = RGB(255, 0, 0)
    fc.Interior.Color = RGB(255, 255, 0)
End Sub

' Side block of the science sheet (W2:Y9): seat shares per branch from the number of passing pupils
Private Sub BuildSideBlock(ByVal ws As Worksheet, ByVal avgHeader As String)
    Dim fnt As String, names As Variant, lo As Variant, hi As Variant, fills As Variant, i As Long, r As Long
    fnt = "Calibri Light"
    ws.Columns(22).ColumnWidth = 6.2
    ws.Columns(23).ColumnWidth = 13.3
    ws.Columns(24).ColumnWidth = 8.6
    ws.Columns(25).ColumnWidth = 9.1

    ws.Range("W2").Value2 = Txt("627 644 62A 640 62D 640 62C 64A 640 645 20 641 64A 20 627 644 62A 640 648 62C 640 64A 640 640 640 647")
    With ws.Range("W2:Y2")
        .Merge
        .Font.Name = fnt: .Font.Size = 22: .Font.Bold = True: .Font.Color = RGB(255, 0, 0)
        .Interior.Color = RGB(255, 255, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ws.Range("W3").Value2 = Txt("62C 20 645 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627")
    With ws.Range("W3:Y3")
        .Merge
        .Font.Name = fnt: .Font.Size = 20: .Font.Bold = True: .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(0, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With

    ws.Range("W4").Value2 = Txt("627 644 646 627 62C 62D 64A 646")
    ws.Range("X4").Formula = "=COUNTIF(Data[" & avgHeader & "],"">=10"")"
    ws.Range("X4:Y4").Merge
    With ws.Range("W4:Y4")
        .Font.Name = fnt: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With

    ws.Range("X5").Value2 = Txt("639 644 649 20 627 644 623 642 644")
    ws.Range("Y5").Value2 = Txt("639 644 649 20 627 644 623 643 62B 631")
    With ws.Range("W5:Y5")
        .Font.Name = fnt: .Font.Size = 14: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With

    names = Array(Txt("631 64A 627 636 64A 627 62A"), Txt("639 644 648 645 20 62A 62C 631 64A 628 64A 629"), Txt("62A 642 646 64A 20 631 64A 627 636 64A"), Txt("62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F"))
    lo = Array(8, 50, 18, 16)
    hi = Array(11, 55, 22, 20)
    fills = Array(RGB(255, 255, 0), RGB(0, 255, 255), RGB(189, 215, 238), RGB(255, 230, 153))
    For i = 0 To 3
        r = 6 + i
        ws.Cells(r, 23).Value2 = names(i)
        ws.Cells(r, 24).Formula = "=(" & lo(i) & "*X4)/100"
        ws.Cells(r, 25).Formula = "=" & hi(i) & "*X4/100"
        With ws.Range(ws.Cells(r, 23), ws.Cells(r, 25))
            .Font.Name = fnt: .Font.Size = 18: .Font.Bold = True
            .Interior.Color = fills(i)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
        ws.Range(ws.Cells(r, 24), ws.Cells(r, 25)).NumberFormat = "0"
    Next i
End Sub

Private Sub GuidancePageSetup(ByVal ws As Worksheet, ByVal isArts As Boolean)
    On Error Resume Next
    With ws.PageSetup
        .PaperSize = xlPaperA4
        If isArts Then
            .Orientation = xlPortrait
        Else
            .Orientation = xlLandscape
            .LeftMargin = Application.InchesToPoints(0.315)
            .RightMargin = Application.InchesToPoints(0.315)
            .TopMargin = Application.InchesToPoints(0.551)
            .BottomMargin = Application.InchesToPoints(0.157)
            .HeaderMargin = Application.InchesToPoints(0.315)
            .FooterMargin = Application.InchesToPoints(0.315)
            .CenterHorizontally = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End If
    End With
    Err.Clear
    On Error GoTo 0
End Sub

' Formats and protects a sheet that received no pupils (empty track)
Private Sub FinishEmptyTrack(ByVal wb As Workbook, ByVal isArts As Boolean)
    Dim ws As Worksheet, lo As ListObject
    If isArts Then
        Set ws = wb.Worksheets("Data01")
        Set lo = GuidanceTable(ws, 15)
    Else
        Set ws = wb.Worksheets("Data")
        Set lo = GuidanceTable(ws, 21)
    End If
    FormatGuidanceBody ws, isArts, 3
    LockAndProtect ws, lo.DataBodyRange
End Sub


' ==============================================================================
' OUTPUT FOLDER
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
