Attribute VB_Name = "modStreamPreferences02"
Option Explicit

Private m_runSucceeded As Boolean
Private m_silent As Boolean
Private m_lastReport As String

' ============================================================
' Preferences 02: Tecno -> Data O:P; Litr -> Data01 K:L, in the guidance workbook that
' modTransfer02 keeps in memory (never opened from disk, never saved here on success).
' If this step cannot complete, a copy of the workbook is saved in the output folder.
' Strict 1-to-1 Matching Rules:
' 1. Exact 1-to-1 match with empty wish cell(s):
'    -> Process CONTINUES, data is transferred, and empty wishes are reported.
' 2. No match (unmatched row in the wishes file OR unmatched pupil in the guidance workbook)
'    OR Multi-row match (duplicate/ambiguous rows in the wishes file OR in the guidance workbook)
'    OR invalid birth dates / repeated non-empty wishes:
'    -> Process STOPS (Preferences02Succeeded = False), target is not modified,
'       and all mismatches/duplicates are announced in the report.
' ============================================================

' The wishes file is found INSIDE Source02 by its STRUCTURE, never by its name:
' sheets Tecno and/or Litr, headers in row 2, B name / C birth date / D wish 1 / E wish 2.
Private Const SRC_FOLDER As String = "Source02"
Private Const WISH_SHEET_SCIENCE As String = "Tecno"
Private Const WISH_SHEET_ARTS As String = "Litr"
Private Const W_FIRST_ROW As Long = 3     ' headers in row 2
Private Const W_COL_NAME As Long = 2      ' B
Private Const W_COL_BIRTH As Long = 3     ' C
Private Const W_COL_W1 As Long = 4        ' D
Private Const W_COL_W2 As Long = 5        ' E

Private Const DST_FILE As String = "Term02Guidance.xlsx"
Private Const C_NAME As Long = 2
Private Const C_BIRTH As Long = 3

Private Type WishesBatch
    Wishes As Variant
    WriteRows As Variant
    Matched As Long
    EmptyWishCount As Long
    UnmatchedSourceCount As Long
    UnmatchedTargetCount As Long
    MultiSourceCount As Long
    MultiTargetCount As Long
    InvalidCount As Long
    HasFatalStop As Boolean
    TrackActive As Boolean
    IsArts As Boolean
    Message As String
    MissingSource As Boolean
End Type

Public Sub SetSilentMode(ByVal silent As Boolean)
    m_silent = silent
End Sub

Public Function LastReport() As String
    LastReport = m_lastReport
End Function

Public Function Preferences02Succeeded() As Boolean
    Preferences02Succeeded = m_runSucceeded
End Function

Public Sub Run_StreamPreferences02()
    Dim stateSaved As Boolean, wasSilent As Boolean
    Dim scienceTouched As Boolean, artsTouched As Boolean
    Dim wbW As Workbook, wbDst As Workbook, wOpened As Boolean, copyPath As String
    Dim science As WishesBatch, arts As WishesBatch
    Dim stopIssues As Object, emptyNotes As Object
    Dim ws As Worksheet, lo As ListObject, failed As Boolean, errText As String, msg As String
    Dim hasStop As Boolean, problem As String, wishPath As String

    wasSilent = m_silent
    m_runSucceeded = False
    m_lastReport = ""

    On Error GoTo EH
    Application.EnableCancelKey = xlErrorHandler

    stateSaved = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set wbDst = modTransfer02.GuidanceWorkbook()
    If wbDst Is Nothing Then Err.Raise vbObjectError + 85, , "No guidance workbook is open. Run the data transfer first (Term 02 Orientation button)."
    problem = FindWishesFile(wishPath)
    If Len(problem) > 0 Then Err.Raise vbObjectError + 83, , problem
    Set wbW = OpenWb(wishPath, True, wOpened)

    Set stopIssues = CreateObject("Scripting.Dictionary")
    Set emptyNotes = CreateObject("Scripting.Dictionary")

    PrepareWishes wbW, wbDst, False, science, stopIssues, emptyNotes
    PrepareWishes wbW, wbDst, True, arts, stopIssues, emptyNotes

    If Not science.TrackActive And Not arts.TrackActive Then
        Err.Raise vbObjectError + 84, , "No active student records found to match between the wishes file and the guidance workbook."
    End If

    hasStop = science.HasFatalStop Or arts.HasFatalStop Or (stopIssues.Count > 0)

    msg = "Wishes file: " & Mid$(wishPath, InStrRev(wishPath, Application.PathSeparator) + 1) & vbLf & science.Message & vbLf & arts.Message
    If emptyNotes.Count > 0 Then
        msg = AppendDetailsList(emptyNotes, msg, "Matched rows with empty wishes (work continues): " & emptyNotes.Count)
    End If

    If hasStop Then
        msg = AppendDetailsList(stopIssues, msg, _
              "STOPPED - Unmatched or multi-row matches detected (" & stopIssues.Count & " issue row(s)):")
        m_lastReport = msg
        Err.Raise vbObjectError + 90, , msg
    End If

    If science.Matched > 0 Then
        scienceTouched = True
        CheckTrackWrite wbDst, RequireWorksheet(wbDst, "Data")
    End If
    If arts.Matched > 0 Then
        artsTouched = True
        CheckTrackWrite wbDst, RequireWorksheet(wbDst, "Data01")
    End If

    WriteWishes wbDst, science
    WriteWishes wbDst, arts

    m_runSucceeded = True
    ' Run on its own (not part of the orientation sequence): keep a copy so the work is not lost
    If Not wasSilent Then msg = msg & vbLf & vbLf & "Saved to: " & modTransfer02.SaveGuidanceCopy()
    m_lastReport = msg

Done:
    On Error Resume Next
    If wOpened And Not wbW Is Nothing Then wbW.Close SaveChanges:=False
    If failed And Not wbDst Is Nothing Then
        For Each ws In wbDst.Worksheets
            If (ws.Name = "Data" And scienceTouched) Or (ws.Name = "Data01" And artsTouched) Then
                Set lo = Nothing
                If ws.Name = "Data01" Then Set lo = GuidanceTable(ws, 15) Else Set lo = GuidanceTable(ws, 21)
                If Not lo Is Nothing Then LockAndProtect ws, lo.DataBodyRange
            End If
        Next ws
        ' The step did not complete: keep what was transferred so far as a saved copy
        copyPath = ""
        copyPath = modTransfer02.SaveGuidanceCopy()
        If Len(copyPath) > 0 Then errText = errText & vbLf & vbLf & "A copy of the guidance workbook (without the wishes) was saved: " & copyPath
    End If
    If stateSaved Then
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If
    Application.EnableCancelKey = xlInterrupt
    m_silent = False
    On Error GoTo 0

    If failed Then
        m_lastReport = errText
        If Not wasSilent Then
            MsgBox "Wishes import stopped:" & vbLf & errText, vbCritical, "Stream preferences 02"
        End If
    Else
        If Not wasSilent Then
            MsgBox msg, IIf(emptyNotes.Count > 0, vbExclamation, vbInformation), "Stream preferences 02"
        End If
    End If
    Exit Sub

EH:
    m_runSucceeded = False
    failed = True
    If Err.Number = 18 Then
        errText = "Execution was canceled by the user."
    Else
        errText = Err.Description
    End If
    Resume Done
End Sub

Private Sub PrepareWishes(ByVal wbW As Workbook, ByVal wbDst As Workbook, ByVal isArts As Boolean, _
                          ByRef batch As WishesBatch, ByVal stopIssues As Object, ByVal emptyNotes As Object)
    Dim wsW As Worksheet, wsDst As Worksheet, lo As ListObject, body As Range
    Dim sourceSheet As String, targetSheet As String, wishColumn As Long, width As Long
    Dim tv As Variant, wv As Variant, wishOut As Variant, writeRows() As Boolean
    Dim dFull As Object, dNameCnt As Object, dNameRow As Object, dSource As Object, dWritten As Object
    Dim sourceKey As String, tdk() As String
    Dim nT As Long, i As Long, r As Long, lastRow As Long, rowT As Long
    Dim nk As String, dk As String, fk As String
    Dim born As Date, location As String, firstSourceRow As Long, firstSourceIndex As Long
    Dim firstTargetRow As Long
    Dim sourceNames() As String, sourceDates() As String, sourceOKRows() As Boolean
    Dim sourceTargets() As Long, badTargetDate() As Boolean, targetDup() As Boolean
    Dim nSource As Long, hasTargetPupils As Boolean, hasSourceRows As Boolean
    Dim w1Clean As Variant, w2Clean As Variant

    batch.IsArts = isArts
    If isArts Then
        sourceSheet = WISH_SHEET_ARTS: targetSheet = "Data01": wishColumn = 11: width = 15
    Else
        sourceSheet = WISH_SHEET_SCIENCE: targetSheet = "Data": wishColumn = 15: width = 21
    End If

    Set wsW = FindWorksheet(wbW, sourceSheet)
    Set wsDst = FindWorksheet(wbDst, targetSheet)

    If Not wsDst Is Nothing Then
        Set lo = GuidanceTable(wsDst, width)
        wsDst.Calculate
        hasTargetPupils = Not GuidanceTableIsEmpty(lo)
    End If

    If Not wsW Is Nothing Then
        wsW.Calculate
        lastRow = wsW.Cells(wsW.Rows.Count, W_COL_NAME).End(xlUp).Row
        If lastRow >= W_FIRST_ROW Then
            wv = wsW.Range(wsW.Cells(W_FIRST_ROW, 1), wsW.Cells(lastRow, W_COL_W2)).Value2
            nSource = UBound(wv, 1)
            For r = 1 To nSource
                If Not BlankValue(wv(r, W_COL_NAME)) Or Not BlankValue(wv(r, W_COL_BIRTH)) Or _
                   Not BlankValue(wv(r, W_COL_W1)) Or Not BlankValue(wv(r, W_COL_W2)) Then
                    hasSourceRows = True
                    Exit For
                End If
            Next r
        End If
    End If

    ' If neither target nor source has students for this track, skip cleanly
    If Not hasTargetPupils And Not hasSourceRows Then
        batch.TrackActive = False
        batch.Message = targetSheet & ": no pupils in track; skipped."
        Exit Sub
    End If

    batch.TrackActive = True

    ' Target has pupils, but wishes sheet is missing or empty -> fatal stop
    If hasTargetPupils And (wsW Is Nothing Or Not hasSourceRows) Then
        batch.MissingSource = True
        batch.HasFatalStop = True
        AddValidationIssue stopIssues, targetSheet, "UNMATCHED_DST(all rows - sheet '" & sourceSheet & "' missing/empty)"
        batch.Message = targetSheet & ": wishes sheet '" & sourceSheet & "' missing or empty."
        Exit Sub
    End If

    ' Source has rows, but target sheet is missing or empty -> fatal stop
    If hasSourceRows And (wsDst Is Nothing Or Not hasTargetPupils) Then
        batch.HasFatalStop = True
        AddValidationIssue stopIssues, sourceSheet, "UNMATCHED_SRC(all rows - target '" & targetSheet & "' missing/empty)"
        batch.Message = targetSheet & ": target sheet empty while '" & sourceSheet & "' has rows."
        Exit Sub
    End If

    Set body = lo.DataBodyRange
    tv = body.Value2
    nT = UBound(tv, 1)
    ReDim writeRows(1 To nT)
    wishOut = body.Columns(wishColumn).Resize(, 2).Value2

    Set dSource = CreateObject("Scripting.Dictionary")
    Set dWritten = CreateObject("Scripting.Dictionary")
    Set dFull = CreateObject("Scripting.Dictionary")
    Set dNameCnt = CreateObject("Scripting.Dictionary")
    Set dNameRow = CreateObject("Scripting.Dictionary")
    ReDim tdk(1 To nT)
    ReDim badTargetDate(1 To nT)
    ReDim targetDup(1 To nT)

    ' ---------- 1. Index and validate target students (Term02Guidance.xlsx) ----------
    For i = 1 To nT
        location = wsDst.Name & "!" & (body.Row + i - 1)
        nk = NameKey(tv(i, C_NAME))
        If Len(nk) = 0 Then
            batch.InvalidCount = batch.InvalidCount + 1
            batch.HasFatalStop = True
            AddValidationIssue stopIssues, location, "MISSING_NAME_DST"
        Else
            If Not BlankValue(tv(i, C_BIRTH)) Then
                If TryBirthDate(tv(i, C_BIRTH), wbDst.Date1904, born) Then
                    tdk(i) = DateKey(tv(i, C_BIRTH), wbDst.Date1904)
                Else
                    AddValidationIssue stopIssues, location, "BAD_DATE_DST(C)"
                    badTargetDate(i) = True
                    batch.InvalidCount = batch.InvalidCount + 1
                    batch.HasFatalStop = True
                End If
            End If

            If dNameCnt.Exists(nk) Then
                dNameCnt(nk) = dNameCnt(nk) + 1
                ' If either target row with the same name lacks a birth date, it is ambiguous in target
                firstTargetRow = CLng(dNameRow(nk))
                If (Len(tdk(i)) = 0 And Not badTargetDate(i)) Or (Len(tdk(firstTargetRow)) = 0 And Not badTargetDate(firstTargetRow)) Then
                    If Not targetDup(firstTargetRow) Then
                        targetDup(firstTargetRow) = True
                        batch.MultiTargetCount = batch.MultiTargetCount + 1
                    End If
                    If Not targetDup(i) Then
                        targetDup(i) = True
                        batch.MultiTargetCount = batch.MultiTargetCount + 1
                    End If
                    batch.HasFatalStop = True
                    AddValidationIssue stopIssues, wsDst.Name & "!" & (body.Row + firstTargetRow - 1), "MULTI_DST(" & (body.Row + i - 1) & ")"
                    AddValidationIssue stopIssues, location, "MULTI_DST(" & (body.Row + firstTargetRow - 1) & ")"
                End If
            Else
                dNameCnt.Add nk, 1
                dNameRow.Add nk, i
            End If

            If Len(tdk(i)) > 0 Then
                fk = nk & "|" & tdk(i)
                If dFull.Exists(fk) Then
                    firstTargetRow = Abs(CLng(dFull(fk)))
                    dFull(fk) = -firstTargetRow            ' duplicate in target
                    If Not targetDup(firstTargetRow) Then
                        targetDup(firstTargetRow) = True
                        batch.MultiTargetCount = batch.MultiTargetCount + 1
                    End If
                    If Not targetDup(i) Then
                        targetDup(i) = True
                        batch.MultiTargetCount = batch.MultiTargetCount + 1
                    End If
                    batch.HasFatalStop = True
                    AddValidationIssue stopIssues, wsDst.Name & "!" & (body.Row + firstTargetRow - 1), "MULTI_DST(" & (body.Row + i - 1) & ")"
                    AddValidationIssue stopIssues, location, "MULTI_DST(" & (body.Row + firstTargetRow - 1) & ")"
                Else
                    dFull.Add fk, i
                End If
            End If
        End If
    Next i

    ' ---------- 2. Validate ALL source records (the wishes file) ----------
    ReDim sourceNames(1 To nSource)
    ReDim sourceDates(1 To nSource)
    ReDim sourceOKRows(1 To nSource)
    ReDim sourceTargets(1 To nSource)

    For r = 1 To nSource
        location = sourceSheet & "!" & (W_FIRST_ROW + r - 1)
        nk = NameKey(wv(r, W_COL_NAME))
        sourceNames(r) = nk

        If Len(nk) = 0 Then
            If Not BlankValue(wv(r, W_COL_BIRTH)) Or Not BlankValue(wv(r, W_COL_W1)) Or Not BlankValue(wv(r, W_COL_W2)) Then
                batch.UnmatchedSourceCount = batch.UnmatchedSourceCount + 1
                batch.HasFatalStop = True
                AddValidationIssue stopIssues, location, "UNMATCHED_SRC(missing name)"
            End If
        Else
            sourceOKRows(r) = True
            dk = ""
            If Not BlankValue(wv(r, W_COL_BIRTH)) Then
                If TryBirthDate(wv(r, W_COL_BIRTH), wbW.Date1904, born) Then
                    dk = DateKey(wv(r, W_COL_BIRTH), wbW.Date1904)
                Else
                    AddValidationIssue stopIssues, location, "BAD_DATE_SRC(C)"
                    sourceOKRows(r) = False
                    batch.InvalidCount = batch.InvalidCount + 1
                    batch.HasFatalStop = True
                End If
            End If
            sourceDates(r) = dk

            If sourceOKRows(r) Then
                sourceKey = nk & "|" & dk
                If dSource.Exists(sourceKey) Then
                    firstSourceRow = CLng(dSource(sourceKey))
                    firstSourceIndex = firstSourceRow - W_FIRST_ROW + 1
                    AddValidationIssue stopIssues, firstSourceRowLabel(sourceSheet, firstSourceRow), "MULTI_SRC(" & (W_FIRST_ROW + r - 1) & ")"
                    AddValidationIssue stopIssues, location, "MULTI_SRC(" & firstSourceRow & ")"
                    sourceOKRows(r) = False
                    sourceOKRows(firstSourceIndex) = False
                    batch.MultiSourceCount = batch.MultiSourceCount + 1
                    batch.HasFatalStop = True
                Else
                    dSource.Add sourceKey, W_FIRST_ROW + r - 1
                End If
            End If

            ' Validate wishes syntax (empty wishes are valid!)
            If IsUnknownWish(wv(r, W_COL_W1), isArts) Then
                AddValidationIssue stopIssues, location, "INVALID_WISH(D)"
                sourceOKRows(r) = False
                batch.InvalidCount = batch.InvalidCount + 1
                batch.HasFatalStop = True
            End If
            If IsUnknownWish(wv(r, W_COL_W2), isArts) Then
                AddValidationIssue stopIssues, location, "INVALID_WISH(E)"
                sourceOKRows(r) = False
                batch.InvalidCount = batch.InvalidCount + 1
                batch.HasFatalStop = True
            End If
            If Not CollectWishPairIssues(CleanWish(wv(r, W_COL_W1), isArts), CleanWish(wv(r, W_COL_W2), isArts), _
                                         location, "D", "E", stopIssues, isArts) Then
                sourceOKRows(r) = False
                batch.InvalidCount = batch.InvalidCount + 1
                batch.HasFatalStop = True
            End If
        End If
    Next r

    ' ---------- 3. Match source records to target records ----------
    For r = 1 To nSource
        If Len(sourceNames(r)) > 0 And sourceOKRows(r) Then
            nk = sourceNames(r)
            dk = sourceDates(r)
            location = sourceSheet & "!" & (W_FIRST_ROW + r - 1)
            rowT = 0

            If Not dNameCnt.Exists(nk) Then
                ' No match in target
                batch.UnmatchedSourceCount = batch.UnmatchedSourceCount + 1
                batch.HasFatalStop = True
                AddValidationIssue stopIssues, location, "UNMATCHED_SRC(" & CStr(wv(r, W_COL_NAME)) & ")"
                sourceOKRows(r) = False
            ElseIf Len(dk) > 0 Then
                fk = nk & "|" & dk
                If dFull.Exists(fk) Then
                    rowT = CLng(dFull(fk))
                    If rowT < 0 Then
                        ' Multiple rows in target matched this student
                        rowT = 0
                        batch.MultiTargetCount = batch.MultiTargetCount + 1
                        batch.HasFatalStop = True
                        AddValidationIssue stopIssues, location, "MULTI_DST(" & CStr(wv(r, W_COL_NAME)) & ")"
                        sourceOKRows(r) = False
                    End If
                ElseIf dNameCnt(nk) = 1 Then
                    If Len(tdk(dNameRow(nk))) = 0 And Not badTargetDate(dNameRow(nk)) Then
                        rowT = dNameRow(nk)
                    Else
                        batch.UnmatchedSourceCount = batch.UnmatchedSourceCount + 1
                        batch.HasFatalStop = True
                        AddValidationIssue stopIssues, location, "UNMATCHED_SRC(date mismatch: " & CStr(wv(r, W_COL_NAME)) & ")"
                        sourceOKRows(r) = False
                    End If
                Else
                    batch.UnmatchedSourceCount = batch.UnmatchedSourceCount + 1
                    batch.HasFatalStop = True
                    AddValidationIssue stopIssues, location, "UNMATCHED_SRC(date mismatch: " & CStr(wv(r, W_COL_NAME)) & ")"
                    sourceOKRows(r) = False
                End If
            Else
                ' Source row has no birth date
                If dNameCnt(nk) = 1 Then
                    rowT = dNameRow(nk)
                Else
                    ' Multiple rows in target share this name
                    batch.MultiTargetCount = batch.MultiTargetCount + 1
                    batch.HasFatalStop = True
                    AddValidationIssue stopIssues, location, "MULTI_DST(ambiguous name: " & CStr(wv(r, W_COL_NAME)) & ")"
                    sourceOKRows(r) = False
                End If
            End If

            If rowT > 0 Then
                If badTargetDate(rowT) Or targetDup(rowT) Then
                    sourceOKRows(r) = False
                    batch.HasFatalStop = True
                ElseIf dWritten.Exists(CStr(rowT)) Then
                    ' More than one row in wishes file matched the same target pupil
                    firstSourceRow = Abs(CLng(dWritten(CStr(rowT))))
                    firstSourceIndex = firstSourceRow - W_FIRST_ROW + 1
                    AddValidationIssue stopIssues, firstSourceRowLabel(sourceSheet, firstSourceRow), "MULTI_SRC(" & (W_FIRST_ROW + r - 1) & ")"
                    AddValidationIssue stopIssues, location, "MULTI_SRC(" & firstSourceRow & ")"
                    sourceOKRows(r) = False
                    sourceOKRows(firstSourceIndex) = False
                    dWritten(CStr(rowT)) = -firstSourceRow
                    batch.MultiSourceCount = batch.MultiSourceCount + 1
                    batch.HasFatalStop = True
                Else
                    dWritten.Add CStr(rowT), W_FIRST_ROW + r - 1
                    sourceTargets(r) = rowT
                End If
            End If
        End If
    Next r

    ' ---------- 4. Check for target students that had NO match in wishes file ----------
    For i = 1 To nT
        If Len(NameKey(tv(i, C_NAME))) > 0 And Not targetDup(i) Then
            If Not dWritten.Exists(CStr(i)) Then
                batch.UnmatchedTargetCount = batch.UnmatchedTargetCount + 1
                batch.HasFatalStop = True
                AddValidationIssue stopIssues, wsDst.Name & "!" & (body.Row + i - 1), _
                                   "UNMATCHED_DST(" & CStr(tv(i, C_NAME)) & ")"
            End If
        End If
    Next i

    ' ---------- 5. Build valid 1-to-1 assignments & track empty wishes ----------
    For r = 1 To nSource
        If Len(sourceNames(r)) > 0 And sourceOKRows(r) And sourceTargets(r) > 0 Then
            rowT = sourceTargets(r)
            w1Clean = CleanWish(wv(r, W_COL_W1), isArts)
            w2Clean = CleanWish(wv(r, W_COL_W2), isArts)
            wishOut(rowT, 1) = w1Clean
            wishOut(rowT, 2) = w2Clean
            writeRows(rowT) = True
            batch.Matched = batch.Matched + 1

            ' Single match with empty wish cell(s): work continues and announces it in report
            If BlankValue(w1Clean) Or BlankValue(w2Clean) Then
                batch.EmptyWishCount = batch.EmptyWishCount + 1
                If BlankValue(w1Clean) And BlankValue(w2Clean) Then
                    AddValidationIssue emptyNotes, wsDst.Name & "!" & (body.Row + rowT - 1), _
                                       CStr(tv(rowT, C_NAME)) & " (Wish 1 & 2 empty)"
                ElseIf BlankValue(w1Clean) Then
                    AddValidationIssue emptyNotes, wsDst.Name & "!" & (body.Row + rowT - 1), _
                                       CStr(tv(rowT, C_NAME)) & " (Wish 1 empty)"
                Else
                    AddValidationIssue emptyNotes, wsDst.Name & "!" & (body.Row + rowT - 1), _
                                       CStr(tv(rowT, C_NAME)) & " (Wish 2 empty)"
                End If
            End If
        End If
    Next r

    batch.Wishes = wishOut
    batch.WriteRows = writeRows

    If batch.HasFatalStop Then
        batch.Message = targetSheet & ": STOPPED (matched: " & batch.Matched & _
                        ", unmatched in wishes: " & batch.UnmatchedSourceCount & _
                        ", unmatched in target: " & batch.UnmatchedTargetCount & _
                        ", multi-match in wishes: " & batch.MultiSourceCount & _
                        ", multi-match in target: " & batch.MultiTargetCount & ")."
    Else
        batch.Message = targetSheet & ": 100% matched (" & batch.Matched & " pupil(s) transferred" & _
                        "; empty wish cells: " & batch.EmptyWishCount & ")."
    End If
End Sub

Private Function firstSourceRowLabel(ByVal sheetName As String, ByVal rowNo As Long) As String
    firstSourceRowLabel = sheetName & "!" & rowNo
End Function

Private Sub WriteWishes(ByVal wb As Workbook, ByRef batch As WishesBatch)
    Dim ws As Worksheet, lo As ListObject, body As Range, rows As Variant, wishes As Variant
    Dim width As Long, firstWish As Long, i As Long, pair(1 To 1, 1 To 2) As Variant
    If batch.Matched = 0 Or batch.HasFatalStop Then Exit Sub
    If batch.IsArts Then
        Set ws = wb.Worksheets("Data01"): width = 15: firstWish = 11
    Else
        Set ws = wb.Worksheets("Data"): width = 21: firstWish = 15
    End If
    Set lo = GuidanceTable(ws, width)
    Set body = lo.DataBodyRange
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    InvalidateOrientation wb, ws
    rows = batch.WriteRows
    wishes = batch.Wishes
    For i = LBound(rows) To UBound(rows)
        If rows(i) Then
            pair(1, 1) = wishes(i, 1): pair(1, 2) = wishes(i, 2)
            ws.Cells(body.Row + i - 1, firstWish).Resize(1, 2).Value2 = pair
        End If
    Next i
    LockAndProtect ws, body
End Sub

' Name key: normalized, lower-case, tokens sorted alphabetically (surname/name order does not matter)
Private Function NameKey(ByVal x As Variant) As String
    Dim s As String, t() As String, i As Long, j As Long, tmp As String, n As Long
    If IsError(x) Then Exit Function
    s = LCase$(Norm(x))
    s = Replace(s, ".", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "_", " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    t = Split(s, " ")
    n = UBound(t)
    For i = 1 To n
        tmp = t(i)
        j = i - 1
        Do While j >= 0
            If StrComp(t(j), tmp, vbBinaryCompare) > 0 Then
                t(j + 1) = t(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        t(j + 1) = tmp
    Next i
    NameKey = Join(t, " ")
End Function

Private Function DateKey(ByVal x As Variant, Optional ByVal date1904 As Boolean = False) As String
    Dim d As Date
    If IsError(x) Then Err.Raise vbObjectError + 44, , "Error value in a birth date."
    If BlankValue(x) Then Exit Function
    If Not TryBirthDate(x, date1904, d) Then
        Err.Raise vbObjectError + 44, , "Invalid birth date: " & CStr(x) & ". No wishes were written."
    End If
    DateKey = Format$(d, "yyyymmdd")
End Function

Private Function Canon(ByVal x As Variant) As Variant
    Dim t As String
    Canon = Empty
    If IsError(x) Then Exit Function
    t = Norm(x)
    If Len(t) = 0 Then Exit Function
    If InStr(t, Norm(ArtsText("f_lang"))) > 0 Then
        Canon = ArtsText("lang")
    ElseIf InStr(t, Norm(ArtsText("f_arts"))) > 0 Or InStr(t, Norm(ArtsText("f_philo"))) > 0 Then
        Canon = ArtsText("arts")
    ElseIf InStr(t, Norm(L("f_tech"))) > 0 Then
        Canon = L("tech")
    ElseIf InStr(t, Norm(L("f_mgmt"))) > 0 Then
        Canon = L("mgmt")
    ElseIf InStr(t, Norm(L("f_exp"))) > 0 Then
        Canon = L("exp")
    ElseIf InStr(t, Norm(L("f_math"))) > 0 Then
        Canon = L("math")
    Else
        Canon = Trim$(CStr(x))
    End If
End Function

' ---------- Common helpers ----------
Private Function Txt(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = 0 To UBound(parts)
        s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    Txt = s
End Function

Private Function L(ByVal key As String) As String
    Select Case key
    Case "math": L = Txt("631 64A 627 636 64A 627 62A")
    Case "tech": L = Txt("62A 642 646 64A 20 631 64A 627 636 64A")
    Case "exp": L = Txt("639 644 648 645 20 62A 62C 631 64A 628 64A 629")
    Case "mgmt": L = Txt("62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F")
    Case "orient": L = Txt("62A 648 62C 64A 647")
    Case "repeat": L = Txt("625 639 627 62F 629 20 627 644 633 646 629")
    Case "yes": L = Txt("646 639 645")
    Case "sheet": L = ArtsText("science_report")
    Case "title": L = Txt("62D 648 635 644 629 20 627 642 62A 631 627 62D 627 62A 20 627 644 62A 648 62C 64A 647")
    Case "c_branch": L = Txt("627 644 634 639 628 629")
    Case "c_w1": L = Txt("627 644 631 63A 628 629 20 30 31")
    Case "c_w2": L = Txt("627 644 631 63A 628 629 20 30 32")
    Case "c_wsum": L = Txt("645 62C 645 648 639 20 627 644 631 63A 628 627 62A")
    Case "c_dec": L = Txt("642 631 627 631 627 62A 20 627 644 62A 648 62C 64A 647")
    Case "c_met": L = Txt("62A 644 628 64A 629 20 627 644 631 63A 628 629 20 30 31")
    Case "c_pct": L = Txt("627 644 646 633 628 629 20 645 646 20 627 644 646 627 62C 62D 64A 646")
    Case "total": L = Txt("627 644 645 62C 645 648 639")
    Case "item": L = Txt("627 644 628 64A 627 646")
    Case "count": L = Txt("627 644 639 62F 62F")
    Case "n_students": L = Txt("639 62F 62F 20 627 644 62A 644 627 645 64A 630")
    Case "n_pass": L = Txt("639 62F 62F 20 627 644 646 627 62C 62D 64A 646")
    Case "n_rep": L = Txt("639 62F 62F 20 642 631 627 631 627 62A 20 625 639 627 62F 629 20 627 644 633 646 629")
    Case "n_ori": L = Txt("639 62F 62F 20 642 631 627 631 627 62A 20 62A 648 62C 64A 647")
    Case "f_math": L = Txt("631 64A 627 636")
    Case "f_tech": L = Txt("62A 642 646 64A")
    Case "f_mgmt": L = Txt("62A 633 64A 64A 631")
    Case "f_exp": L = Txt("639 644 648 645")
    End Select
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
    If Dir(fullPath) = "" Then Err.Raise vbObjectError + 1, , "File not found: " & fullPath
    Set OpenWb = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=asReadOnly)
    opened = True
End Function

' ==============================================================================
' WISHES FILE DISCOVERY (by structure, never by name)
' ==============================================================================
' Returns "" and sets fullPath when exactly ONE file in Source02 has the wishes
' structure; otherwise returns the text to show (none found / several found).
Public Function FindWishesFile(ByRef fullPath As String) As String
    Dim folder As String, f As String, ext As String, names As Collection, i As Long
    Dim wb As Workbook, opened As Boolean, state As Long
    Dim hits As String, near As String, nHits As Long, firstHit As String

    fullPath = ""
    If Len(ThisWorkbook.Path) = 0 Then
        FindWishesFile = "Save the tool workbook before running this procedure."
        Exit Function
    End If
    folder = ThisWorkbook.Path & Application.PathSeparator & SRC_FOLDER & Application.PathSeparator
    If Not WishFolderExists(folder) Then
        FindWishesFile = "The folder '" & SRC_FOLDER & "' was not found next to the tool:" & vbLf & folder
        Exit Function
    End If

    Set names = New Collection
    f = Dir(folder & "*.xls*")
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then
            ext = LCase$(Mid$(f, InStrRev(f, ".") + 1))
            If ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb" Then names.Add f
        End If
        f = Dir
    Loop

    For i = 1 To names.Count
        f = CStr(names(i))
        If StrComp(folder & f, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            Set wb = WishOpenForInspection(folder & f, opened)
            If Not wb Is Nothing Then
                state = WishesLayoutState(wb)
                If state = 2 Then
                    nHits = nHits + 1
                    hits = hits & vbLf & "  - " & f
                    If nHits = 1 Then firstHit = folder & f
                ElseIf state = 1 Then
                    near = near & vbLf & "  - " & f
                End If
                If opened Then wb.Close SaveChanges:=False
            End If
        End If
    Next i

    If nHits = 1 Then
        fullPath = firstHit
    ElseIf nHits = 0 Then
        FindWishesFile = "No wishes file was found in '" & SRC_FOLDER & "'." & vbLf & _
            "It is recognised by its structure (sheets " & WISH_SHEET_SCIENCE & " / " & WISH_SHEET_ARTS & _
            ", headers in row 2: name, birth date, wish 1, wish 2), not by its name."
        If Len(near) > 0 Then FindWishesFile = FindWishesFile & vbLf & "Files with these sheets but different headers:" & near
    Else
        FindWishesFile = "More than one wishes file was found in '" & SRC_FOLDER & "' - keep only one:" & hits
    End If
End Function

' 0 = not a wishes file, 1 = has a wishes sheet but the headers do not match, 2 = wishes file
Private Function WishesLayoutState(ByVal wb As Workbook) As Long
    Dim ws As Worksheet, nm As Variant, found As Long, good As Long
    For Each nm In Array(WISH_SHEET_SCIENCE, WISH_SHEET_ARTS)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            found = found + 1
            If WishHeadersOk(ws) Then good = good + 1
        End If
    Next nm
    If found = 0 Then
        WishesLayoutState = 0
    ElseIf good = found Then
        WishesLayoutState = 2
    Else
        WishesLayoutState = 1
    End If
End Function

' Row 2 headers: B name, C birth date, D wish 1, E wish 2
Private Function WishHeadersOk(ByVal ws As Worksheet) As Boolean
    WishHeadersOk = WishCellHas(ws, 2, "627 633 645") And WishCellHas(ws, 3, "645 64A 644 627 62F") _
                And WishCellHas(ws, 4, "631 63A 628") And WishCellHas(ws, 5, "631 63A 628")
End Function

Private Function WishCellHas(ByVal ws As Worksheet, ByVal col As Long, ByVal wordCodes As String) As Boolean
    Dim v As Variant, w As String, p As Variant
    On Error Resume Next
    v = ws.Cells(W_FIRST_ROW - 1, col).MergeArea.Cells(1, 1).Value2
    On Error GoTo 0
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    For Each p In Split(wordCodes, " ")
        If Len(p) > 0 Then w = w & ChrW(CLng("&H" & p))
    Next p
    WishCellHas = (Len(w) > 0 And InStr(Norm(v), Norm(w)) > 0)
End Function

Private Function WishOpenForInspection(ByVal fullPath As String, ByRef opened As Boolean) As Workbook
    Dim wb As Workbook
    opened = False
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set WishOpenForInspection = wb
            Exit Function
        End If
    Next wb
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True, AddToMRU:=False)
    If Err.Number <> 0 Then
        Err.Clear
        Set wb = Nothing
    End If
    On Error GoTo 0
    If Not wb Is Nothing Then
        opened = True
        Set WishOpenForInspection = wb
    End If
End Function

Private Function WishFolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    WishFolderExists = ((GetAttr(p) And vbDirectory) <> 0)
    If Err.Number <> 0 Then WishFolderExists = False: Err.Clear
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

Private Function KnownWish(ByVal x As Variant, Optional ByVal isArts As Boolean = False) As Boolean
    Dim t As String, k As Long, labels As Variant
    If IsError(x) Then Exit Function
    If BlankValue(x) Then
        KnownWish = True
        Exit Function
    End If
    t = Norm(x)
    If isArts Then
        labels = Array(ArtsText("arts"), ArtsText("lang"))
    Else
        labels = Array(L("math"), L("tech"), L("exp"), L("mgmt"))
    End If
    For k = LBound(labels) To UBound(labels)
        If t = Norm(labels(k)) Then
            KnownWish = True
            Exit Function
        End If
    Next k
End Function

Private Function CleanWish(ByVal x As Variant, Optional ByVal isArts As Boolean = False) As Variant
    Dim c As Variant
    CleanWish = Empty
    If IsError(x) Then Exit Function
    c = Canon(x)
    If KnownWish(c, isArts) Then CleanWish = c
End Function

Private Function IsUnknownWish(ByVal x As Variant, Optional ByVal isArts As Boolean = False) As Boolean
    If IsError(x) Then
        IsUnknownWish = True
    Else
        IsUnknownWish = Not KnownWish(Canon(x), isArts)
    End If
End Function

Private Function CollectWishPairIssues(ByVal w1 As Variant, ByVal w2 As Variant, _
                                       ByVal location As String, ByVal col1 As String, _
                                       ByVal col2 As String, ByVal problems As Object, Optional ByVal isArts As Boolean = False) As Boolean
    Dim one As Variant, two As Variant
    one = Canon(w1): two = Canon(w2)
    CollectWishPairIssues = True
    If IsError(w1) Or Not KnownWish(one, isArts) Then
        AddValidationIssue problems, location, col1
        CollectWishPairIssues = False
    End If
    If IsError(w2) Or Not KnownWish(two, isArts) Then
        AddValidationIssue problems, location, col2
        CollectWishPairIssues = False
    End If
    If Not BlankValue(one) And Not BlankValue(two) Then
        If Norm(one) = Norm(two) Then
            AddValidationIssue problems, location, col1 & "=" & col2
            CollectWishPairIssues = False
        End If
    End If
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

Private Sub AddValidationIssue(ByVal problems As Object, ByVal location As String, ByVal detail As String)
    If problems.Exists(location) Then
        If InStr(1, " " & problems(location) & " ", " " & detail & " ", vbBinaryCompare) = 0 Then
            problems(location) = problems(location) & " | " & detail
        End If
    Else
        problems.Add location, detail
    End If
End Sub

Private Function AppendDetailsList(ByVal items As Object, ByVal baseMsg As String, ByVal sectionHeader As String) As String
    Const MAX_ROWS As Long = 15
    Const MAX_LINE As Long = 65
    Dim key As Variant, shown As Long, line As String, msg As String
    msg = baseMsg
    If items.Count > 0 Then
        msg = msg & vbLf & sectionHeader
        For Each key In items.Keys
            If shown >= MAX_ROWS Then Exit For
            line = "  - " & CStr(key) & ": " & items(key)
            If Len(line) > MAX_LINE Then line = Left$(line, MAX_LINE - 3) & "..."
            msg = msg & vbLf & line
            shown = shown + 1
        Next key
        If items.Count > shown Then
            msg = msg & vbLf & "  ... and " & (items.Count - shown) & " more row(s)."
        End If
    End If
    AppendDetailsList = msg
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
    Case "general_report": ArtsText = Txt("62D 648 635 644 629 20 639 627 645 629")
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

Private Function GuidanceTableIsEmpty(ByVal lo As ListObject) As Boolean
    Dim values As Variant, r As Long, c As Long
    GuidanceTableIsEmpty = True
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            If IsError(values(r, c)) Then
                GuidanceTableIsEmpty = False
                Exit Function
            End If
            If Not BlankValue(values(r, c)) Then
                GuidanceTableIsEmpty = False
                Exit Function
            End If
        Next c
    Next r
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