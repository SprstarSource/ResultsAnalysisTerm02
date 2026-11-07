Attribute VB_Name = "modCompare"
Option Explicit

' \Source02\ (raw files, collected in memory by the embedded engine - no Term02Results.xlsx)
' -> a NEW Term02Comparison.xlsx in the output folder next to this tool workbook.
' Every worksheet is processed, including filtered and hidden source rows/sheets.
' A:E are copied as values with their original number formats and date system.
' F:G hold the two term averages; H is their difference rounded to two decimals.
' High and regressed tables only; neutral rows are excluded.
' Aggregate lists and a general report are created across all source sheets.
' Source files and existing comparison exports are never overwritten.
' Every table is ordered by the difference, largest first (+5, +2 ... / -0.97, -1.20 ...);
' the table filters stay available (e.g. to filter one class). Columns = AutoFit.
' Every sheet is protected (no password; filtering allowed).

Private Const OUTPUT_FILE As String = "Term02Comparison.xlsx"
Private Const HEADER_ROW As Long = 2
Private Const FIRST_DATA_ROW As Long = 3
Private Const MAX_ISSUES As Long = 20
Private Const HIGH_LIMIT As Double = 2
Private Const LOW_LIMIT As Double = -0.95

Private Const SOURCE_FOLDER_NAME As String = "Source02"

' ---- Collection engine (embedded copy of the collector engine of Data02 / ResultFlow02 - keep identical) ----
Private Const DATA_START_ROW       As Long = 6      ' header row in the source sheet
Private Const TRAILING_ROWS_SKIP   As Long = 1      ' footer rows after last used row (original logic: -1). Set 0 if none
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

Private Type CompareBatch
    SheetName As String
    Subtitle As String
    Header As Variant
    Values As Variant
    Groups As Variant
    SourceRows As Variant
    Term1Column As Long
    Term2Column As Long
    RowCount As Long
    Counts(1 To 3) As Long
End Type

Public Sub Run_Compare()
    Dim source As Workbook, target As Workbook, ws As Worksheet
    Dim openedSource As Boolean, savedTarget As Boolean, failed As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim stateSaved As Boolean, oldStatus As Variant
    Dim batches() As CompareBatch, issues As Collection, headerIssues As Collection
    Dim issueCount As Long, headerIssueCount As Long, reason As String, errText As String
    Dim folder As String, outDir As String, outputPath As String
    Dim srcCopied As Long, srcSkipped As Long, srcNotes As String
    Dim i As Long, sheetCount As Long, totalPupils As Long
    Dim activeIndices() As Long, activeCount As Long, idx As Long

    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatus = Application.StatusBar
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    folder = ThisWorkbook.Path
    If Len(folder) = 0 Then Err.Raise vbObjectError + 901, "modCompare", Text("save_tool")
    If LCase$(Left$(folder, 4)) = "http" Then Err.Raise vbObjectError + 906, "modCompare", "Use a local or locally synced folder."
    If Not SrcBuildWorkbook(folder & "\" & SOURCE_FOLDER_NAME & "\", source, srcCopied, srcSkipped, srcNotes) Then
        Err.Raise vbObjectError + 904, "modCompare", "No usable source file was found in " & folder & "\" & SOURCE_FOLDER_NAME & vbLf & srcNotes
    End If
    openedSource = True
    sheetCount = source.Worksheets.Count
    If sheetCount = 0 Then Err.Raise vbObjectError + 902, "modCompare", Text("no_sheets")
    ReDim batches(1 To sheetCount)
    Set issues = New Collection
    Set headerIssues = New Collection

    ' Resolve ALL headers before creating an output workbook. Never guess positions.
    For i = 1 To sheetCount
        Set ws = source.Worksheets(i)
        batches(i).SheetName = ws.Name
        reason = ""
        If Not FindTermColumns(ws, batches(i).Term1Column, batches(i).Term2Column, reason) Then
            AddIssue headerIssues, headerIssueCount, ws.Name & "!2:2: " & reason
        End If
    Next i
    If headerIssueCount > 0 Then
        Err.Raise vbObjectError + 903, "modCompare", Text("bad_headers") & IssueSummary(headerIssues, headerIssueCount)
    End If

    For i = 1 To sheetCount
        Application.StatusBar = Text("working") & " " & batches(i).SheetName
        PrepareSheet source.Worksheets(i), batches(i), issues, issueCount
        totalPupils = totalPupils + batches(i).RowCount
    Next i

    Set target = Application.Workbooks.Add(xlWBATWorksheet)
    target.Date1904 = source.Date1904

    ReDim activeIndices(1 To sheetCount)
    For i = 1 To sheetCount
        If batches(i).Counts(1) > 0 Or batches(i).Counts(3) > 0 Then
            activeCount = activeCount + 1
            activeIndices(activeCount) = i
        End If
    Next i
    If activeCount > 0 Then ReDim Preserve activeIndices(1 To activeCount)

    BuildGeneralReport target, batches, activeIndices, activeCount
    BuildAggregatedList source, target, batches, activeIndices, activeCount, 1
    BuildAggregatedList source, target, batches, activeIndices, activeCount, 3
    For i = 1 To activeCount
        idx = activeIndices(i)
        Set ws = target.Worksheets.Add(After:=target.Worksheets(target.Worksheets.Count))
        ws.Name = batches(idx).SheetName
        BuildSheet source.Worksheets(idx), ws, batches(idx), i
    Next i
    target.Worksheets(1).Activate
    ' columns = AutoFit on every sheet (merged titles are ignored by AutoFit), then protect
    For Each ws In target.Worksheets
        ws.UsedRange.EntireColumn.AutoFit
    Next ws
    For Each ws In target.Worksheets
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    Next ws
    outDir = folder & "\" & OutputFolderName()
    If Not EnsureFolder(outDir) Then Err.Raise vbObjectError + 907, "modCompare", "The output folder could not be created: " & outDir
    outputPath = NewOutputPath(outDir)
    target.CheckCompatibility = False
    target.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    savedTarget = True

Finished:
    On Error Resume Next
    If openedSource And Not source Is Nothing Then source.Close SaveChanges:=False
    If failed And Not savedTarget And Not target Is Nothing Then target.Close SaveChanges:=False
    If stateSaved Then
        Application.StatusBar = oldStatus
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    On Error GoTo 0
    If failed Then
        MsgBox Text("failed") & vbLf & errText, vbCritical, "modCompare"
    Else
        errText = Text("done") & vbLf & outputPath & vbLf & _
                  Text("sheets") & ": " & target.Worksheets.Count & vbLf & Text("pupils") & ": " & totalPupils
        If issueCount > 0 Then errText = errText & vbLf & vbLf & Text("skipped") & IssueSummary(issues, issueCount)
        If Len(srcNotes) > 0 Then errText = errText & vbLf & vbLf & "Source notes:" & vbLf & srcNotes
        MsgBox errText, IIf(issueCount > 0, vbExclamation, vbInformation), "modCompare"
    End If
    Exit Sub
Failed:
    failed = True
    errText = Err.Description
    Resume Finished
End Sub

Private Function FindTermColumns(ByVal ws As Worksheet, ByRef first As Long, ByRef second As Long, ByRef reason As String) As Boolean
    Dim lastColumn As Long, c As Long, h As String, term As Long
    Dim firstCount As Long, secondCount As Long
    first = 0: second = 0
    lastColumn = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastColumn)).Calculate
    For c = 6 To lastColumn
        h = Normalize(ws.Cells(HEADER_ROW, c).MergeArea.Cells(1, 1).Value2)
        term = HeaderTerm(h)
        If term = 1 Then
            first = c: firstCount = firstCount + 1
        ElseIf term = 2 Then
            second = c: secondCount = secondCount + 1
        End If
    Next c
    If firstCount <> 1 Or secondCount <> 1 Then
        reason = Text("unique_headers") & " (" & Text("term1") & ": " & firstCount & "; " & Text("term2") & ": " & secondCount & ")"
        Exit Function
    End If
    FindTermColumns = True
End Function

Private Function HeaderTerm(ByVal header As String) As Long
    Dim c As String, digits As String, i As Long, n As Long
    If InStr(header, Normalize(Text("average_word"))) = 0 Then Exit Function
    If InStr(header, Normalize(Text("term_word"))) = 0 Then Exit Function
    For i = 1 To Len(header)
        c = Mid$(header, i, 1)
        If c >= "0" And c <= "9" Then
            digits = digits & c
        ElseIf Len(digits) > 0 Then
            If Len(digits) <= 2 Then
                n = CLng(Val(digits))
                If n = 1 Or n = 2 Then HeaderTerm = n
            End If
            Exit Function
        End If
    Next i
    If Len(digits) > 0 Then
        If Len(digits) <= 2 Then
            n = CLng(Val(digits))
            If n = 1 Or n = 2 Then HeaderTerm = n
        End If
    ElseIf InStr(header, Normalize(Text("second_word"))) > 0 Then
        HeaderTerm = 2
    ElseIf InStr(header, Normalize(Text("first_word"))) > 0 Then
        HeaderTerm = 1
    End If
End Function

Private Sub PrepareSheet(ByVal ws As Worksheet, ByRef batch As CompareBatch, ByVal issues As Collection, ByRef issueCount As Long)
    Dim lastRow As Long, lastColumn As Long, r As Long, c As Long, count As Long, group As Long
    Dim sourceValues As Variant, values() As Variant, groups() As Long, sourceRows() As Long, headers(1 To 1, 1 To 8) As Variant
    Dim first As Double, second As Double, delta As Double, invalid As String
    For c = 1 To 5
        headers(1, c) = ws.Cells(HEADER_ROW, c).MergeArea.Cells(1, 1).Value2
    Next c
    headers(1, 6) = Text("term1")
    headers(1, 7) = Text("term2")
    headers(1, 8) = Text("difference")
    batch.Header = headers
    If Not IsError(ws.Range("H1").MergeArea.Cells(1, 1).Value2) Then batch.Subtitle = CStr(ws.Range("H1").MergeArea.Cells(1, 1).Value2)
    lastColumn = batch.Term1Column
    If batch.Term2Column > lastColumn Then lastColumn = batch.Term2Column
    lastRow = LastStudentRow(ws, lastColumn)
    If lastRow < FIRST_DATA_ROW Then Exit Sub
    ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastRow, 5)).Calculate
    ws.Range(ws.Cells(FIRST_DATA_ROW, batch.Term1Column), ws.Cells(lastRow, batch.Term1Column)).Calculate
    ws.Range(ws.Cells(FIRST_DATA_ROW, batch.Term2Column), ws.Cells(lastRow, batch.Term2Column)).Calculate
    sourceValues = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastRow, lastColumn)).Value2
    ReDim values(1 To UBound(sourceValues, 1), 1 To 8)
    ReDim groups(1 To UBound(sourceValues, 1))
    ReDim sourceRows(1 To UBound(sourceValues, 1))
    For r = 1 To UBound(sourceValues, 1)
        If Not EmptyRecord(sourceValues, r, batch.Term1Column, batch.Term2Column) Then
            invalid = ""
            If IsError(sourceValues(r, 2)) Or Len(Normalize(sourceValues(r, 2))) = 0 Then invalid = "B"
            If Not TryMark(sourceValues(r, batch.Term1Column), first) Then invalid = invalid & " " & ColumnName(batch.Term1Column)
            If Not TryMark(sourceValues(r, batch.Term2Column), second) Then invalid = invalid & " " & ColumnName(batch.Term2Column)
            If Len(invalid) > 0 Then
                AddIssue issues, issueCount, ws.Name & "!" & (r + FIRST_DATA_ROW - 1) & " [" & Trim$(invalid) & "]: " & Text("invalid_row")
            Else
                count = count + 1
                ' Remove binary subtraction noise before rounding to grade precision.
                delta = Sgn(second - first) * Application.WorksheetFunction.Round(Application.WorksheetFunction.Round(Abs(second - first), 10), 2)
                group = DifferenceGroup(delta)
                For c = 2 To 5
                    values(count, c) = sourceValues(r, c)
                Next c
                values(count, 6) = first
                values(count, 7) = second
                values(count, 8) = delta
                groups(count) = group
                sourceRows(count) = r + FIRST_DATA_ROW - 1
                batch.Counts(group) = batch.Counts(group) + 1
            End If
        End If
    Next r
    If count > 1 Then SortRecordsDescending values, groups, sourceRows, count
    batch.RowCount = count
    batch.Values = values
    batch.Groups = groups
    batch.SourceRows = sourceRows
End Sub

Private Function LastStudentRow(ByVal ws As Worksheet, ByVal lastColumn As Long) As Long
    Dim lo As ListObject, c As Long, r As Long
    For Each lo In ws.ListObjects
        If Not lo.HeaderRowRange Is Nothing Then
            If lo.HeaderRowRange.Row = HEADER_ROW And lo.Range.Column = 1 And lo.Range.Columns.Count >= lastColumn Then
                If lo.DataBodyRange Is Nothing Then LastStudentRow = HEADER_ROW Else LastStudentRow = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
                Exit Function
            End If
        End If
    Next lo
    LastStudentRow = HEADER_ROW
    For c = 1 To 5
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > LastStudentRow Then LastStudentRow = r
    Next c
End Function

Private Function EmptyRecord(ByRef values As Variant, ByVal row As Long, ByVal first As Long, ByVal second As Long) As Boolean
    Dim c As Long
    For c = 1 To 5
        If IsError(values(row, c)) Then Exit Function
        If Len(Normalize(values(row, c))) > 0 Then Exit Function
    Next c
    If IsError(values(row, first)) Or IsError(values(row, second)) Then Exit Function
    EmptyRecord = (Len(Normalize(values(row, first))) = 0 And Len(Normalize(values(row, second))) = 0)
End Function

Private Function TryMark(ByVal value As Variant, ByRef mark As Double) As Boolean
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    mark = CDbl(value)
    TryMark = (mark >= 0 And mark <= 20)
End Function

Private Function DifferenceGroup(ByVal difference As Double) As Long
    If difference >= HIGH_LIMIT Then
        DifferenceGroup = 1
    ElseIf difference <= LOW_LIMIT Then
        DifferenceGroup = 3
    Else
        DifferenceGroup = 2
    End If
End Function

Private Sub BuildSheet(ByVal source As Worksheet, ByVal ws As Worksheet, ByRef batch As CompareBatch, ByVal sheetNumber As Long)
    Dim nextRow As Long, lastRow As Long
    ws.DisplayRightToLeft = True
    With ws.Range("A1:H1")
        .Merge
        .Value2 = Text("title") & " - " & batch.SheetName
        .Font.Name = "Dubai": .Font.Size = 16: .Font.Bold = True
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.Rows(1).RowHeight = 36
    With ws.Range("A2:H2")
        .Merge: .Value2 = batch.Subtitle
        .Font.Name = "Dubai": .Font.Size = 14
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.Rows(2).RowHeight = 30
    nextRow = 4
    If batch.Counts(1) > 0 Then nextRow = WriteGroup(source, ws, batch, 1, nextRow, sheetNumber)
    If batch.Counts(3) > 0 Then nextRow = WriteGroup(source, ws, batch, 3, nextRow, sheetNumber)
    lastRow = nextRow - 3
    If lastRow < 3 Then lastRow = 3
    ws.Calculate
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 8)).Address
        .Orientation = xlLandscape: .Zoom = False
        .FitToPagesWide = 1: .FitToPagesTall = False: .CenterHorizontally = True
    End With
End Sub

Private Function WriteGroup(ByVal source As Worksheet, ByVal ws As Worksheet, ByRef batch As CompareBatch, _
                            ByVal group As Long, ByVal titleRow As Long, ByVal sheetNumber As Long) As Long
    Dim firstRow As Long, lastRow As Long, r As Long, c As Long, p As Long
    Dim out() As Variant, values As Variant, groups As Variant, sourceRows As Variant
    Dim lo As ListObject, area As Range, shade As Long, tableRows As Long
    If group = 1 Then shade = RGB(39, 139, 82) Else shade = RGB(190, 55, 55)
    With ws.Range(ws.Cells(titleRow, 1), ws.Cells(titleRow, 8))
        .Merge: .Value2 = Text("group" & CStr(group))
        .Font.Name = "Dubai": .Font.Size = 14: .Font.Bold = True
        .Interior.Color = shade: .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(titleRow).RowHeight = 28
    firstRow = titleRow + 2: tableRows = batch.Counts(group)
    If tableRows = 0 Then
        WriteGroup = titleRow
        Exit Function
    End If
    lastRow = firstRow + tableRows - 1
    If lastRow > ws.Rows.Count Then Err.Raise vbObjectError + 905, "modCompare", Text("too_many_rows") & " " & ws.Name
    Set area = ws.Range(ws.Cells(titleRow + 1, 1), ws.Cells(lastRow, 8))
    area.Font.Name = "Dubai": area.Font.Size = 14: area.Font.Bold = False
    area.HorizontalAlignment = xlCenter: area.VerticalAlignment = xlCenter
    area.ReadingOrder = xlRTL: area.Rows.RowHeight = 24
    area.Borders.LineStyle = xlContinuous: area.Borders.Weight = xlThin
    ws.Cells(titleRow + 1, 1).Resize(1, 8).Value2 = batch.Header
    ReDim out(1 To tableRows, 1 To 8)
    values = batch.Values: groups = batch.Groups: sourceRows = batch.SourceRows
    For r = 1 To batch.RowCount
        If groups(r) = group Then
            p = p + 1: out(p, 1) = p
            For c = 2 To 8: out(p, c) = values(r, c): Next c
        End If
    Next r
    ws.Cells(firstRow, 1).Resize(p, 8).Value2 = out
    p = 0
    For r = 1 To batch.RowCount
        If groups(r) = group Then
            For c = 2 To 5
                ws.Cells(firstRow + p, c).NumberFormat = source.Cells(CLng(sourceRows(r)), c).NumberFormat
            Next c
            p = p + 1
        End If
    Next r
    ws.Range(ws.Cells(firstRow, 6), ws.Cells(lastRow, 7)).NumberFormat = "0.00"
    ws.Range(ws.Cells(firstRow, 8), ws.Cells(lastRow, 8)).NumberFormat = "+0.00;-0.00;0.00"
    With ws.Range(ws.Cells(firstRow, 8), ws.Cells(lastRow, 8))
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        If group = 1 Then .Interior.Color = RGB(39, 139, 82) Else .Interior.Color = RGB(190, 55, 55)
    End With
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=area, XlListObjectHasHeaders:=xlYes)
    lo.Name = "Cmp_" & CStr(sheetNumber) & "_" & CStr(group)
    lo.TableStyle = "": lo.ShowTotals = False
    lo.ShowTableStyleRowStripes = False: lo.ShowTableStyleColumnStripes = False: lo.ShowAutoFilter = True
    With lo.HeaderRowRange
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .WrapText = True: .RowHeight = 36
    End With
    WriteGroup = lastRow + 3
End Function

Private Sub BuildAggregatedList(ByVal source As Workbook, ByVal target As Workbook, ByRef batches() As CompareBatch, _
                                ByRef activeIndices() As Long, ByVal activeCount As Long, ByVal group As Long)
    Dim ws As Worksheet, original As Worksheet, lo As ListObject, area As Range
    Dim total As Long, i As Long, r As Long, c As Long, p As Long, lastRow As Long, idx As Long
    Dim headers(1 To 1, 1 To 9) As Variant, out() As Variant, baseHeaders As Variant
    Dim values As Variant, groups As Variant, sourceRows As Variant, sheetNumber As Long
    Dim entBatch() As Long, entRow() As Long, entDelta() As Double, order() As Long, k As Long, curIdx As Long
    For i = 1 To activeCount
        idx = activeIndices(i): total = total + batches(idx).Counts(group)
    Next i
    If total = 0 Then Exit Sub
    If group = 1 Then
        Set ws = target.Worksheets.Add(After:=target.Worksheets(1))
        ws.Name = Text("high_list")
    Else
        Set ws = target.Worksheets.Add(After:=target.Worksheets(target.Worksheets.Count))
        ws.Name = Text("low_list")
    End If
    If activeCount > 0 Then baseHeaders = batches(activeIndices(1)).Header Else baseHeaders = Empty
    If IsEmpty(baseHeaders) Then
        For c = 1 To 8: headers(1, c) = "": Next c
    Else
        For c = 1 To 8: headers(1, c) = baseHeaders(1, c): Next c
    End If
    headers(1, 9) = Text("class")
    ws.DisplayRightToLeft = True
    With ws.Range("A1:I1")
        .Merge: If group = 1 Then .Value2 = Text("high_list") Else .Value2 = Text("low_list")
        .Font.Name = "Dubai": .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 36
    lastRow = 3 + total: If total = 0 Then lastRow = 3
    Set area = ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 9))
    area.Font.Name = "Dubai": area.Font.Size = 14: area.Font.Bold = False
    area.HorizontalAlignment = xlCenter: area.VerticalAlignment = xlCenter
    area.ReadingOrder = xlRTL: area.Rows.RowHeight = 30
    area.Borders.LineStyle = xlContinuous: area.Borders.Weight = xlThin
    ws.Range("A3:I3").Value2 = headers
    If total > 0 Then
        ReDim out(1 To total, 1 To 9)
        ReDim entBatch(1 To total): ReDim entRow(1 To total)
        ReDim entDelta(1 To total): ReDim order(1 To total)
        For i = 1 To activeCount
            idx = activeIndices(i)
            If batches(idx).Counts(group) > 0 Then
                values = batches(idx).Values: groups = batches(idx).Groups
                For r = 1 To batches(idx).RowCount
                    If groups(r) = group Then
                        p = p + 1
                        entBatch(p) = idx: entRow(p) = r: entDelta(p) = values(r, 8)
                    End If
                Next r
            End If
        Next i
        SortIndexDescending entDelta, order, total
        curIdx = 0
        For p = 1 To total
            k = order(p)
            If entBatch(k) <> curIdx Then
                curIdx = entBatch(k): values = batches(curIdx).Values
            End If
            out(p, 1) = p
            For c = 2 To 8: out(p, c) = values(entRow(k), c): Next c
            out(p, 9) = batches(curIdx).Subtitle
        Next p
        ws.Range("A4").Resize(total, 9).Value2 = out
        curIdx = 0
        For p = 1 To total
            k = order(p)
            If entBatch(k) <> curIdx Then
                curIdx = entBatch(k)
                sourceRows = batches(curIdx).SourceRows
                Set original = source.Worksheets(batches(curIdx).SheetName)
            End If
            For c = 2 To 5: ws.Cells(3 + p, c).NumberFormat = original.Cells(CLng(sourceRows(entRow(k))), c).NumberFormat: Next c
        Next p
        ws.Range("F4:G" & lastRow).NumberFormat = "0.00"
        ws.Range("H4:H" & lastRow).NumberFormat = "+0.00;-0.00;0.00"
        With ws.Range("H4:H" & lastRow)
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            If group = 1 Then .Interior.Color = RGB(39, 139, 82) Else .Interior.Color = RGB(190, 55, 55)
        End With
    End If
    ws.Range("I4:I" & lastRow).WrapText = True
    If total > 0 Then
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=area, XlListObjectHasHeaders:=xlYes)
        If group = 1 Then lo.Name = "Cmp_HighList" Else lo.Name = "Cmp_LowList"
        lo.TableStyle = "": lo.ShowTotals = False
        lo.ShowTableStyleRowStripes = False: lo.ShowTableStyleColumnStripes = False: lo.ShowAutoFilter = True
        With lo.HeaderRowRange
            .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True: .WrapText = True: .RowHeight = 36
        End With
    Else
        ws.Range("A3:I3").Interior.Color = RGB(28, 55, 103)
        ws.Range("A3:I3").Font.Color = RGB(255, 255, 255)
        ws.Range("A3:I3").Font.Bold = True
    End If
    With ws.PageSetup
        .PrintArea = "$A$1:$I$" & lastRow: .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterHorizontally = True: .PrintTitleRows = "$1:$3"
    End With
End Sub

Private Sub BuildGeneralReport(ByVal target As Workbook, ByRef batches() As CompareBatch, _
                               ByRef activeIndices() As Long, ByVal activeCount As Long)
    Dim ws As Worksheet, lo As ListObject, area As Range
    Dim headers(1 To 1, 1 To 3) As Variant, out() As Variant
    Dim i As Long, p As Long, idx As Long, lastRow As Long
    Set ws = target.Worksheets(1): ws.Name = Text("general_report")
    ws.DisplayRightToLeft = True
    With ws.Range("A1:C1")
        .Merge: .Value2 = Text("general_report")
        .Font.Name = "Dubai": .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 36
    headers(1, 1) = Text("pedagogical_group"): headers(1, 2) = Text("high_count"): headers(1, 3) = Text("low_count")
    lastRow = 3 + activeCount: If activeCount = 0 Then lastRow = 3
    Set area = ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 3))
    area.Font.Name = "Dubai": area.Font.Size = 14
    area.HorizontalAlignment = xlCenter: area.VerticalAlignment = xlCenter
    area.ReadingOrder = xlRTL: area.Rows.RowHeight = 30
    area.Borders.LineStyle = xlContinuous: area.Borders.Weight = xlThin
    ws.Range("A3:C3").Value2 = headers
    If activeCount > 0 Then
        ReDim out(1 To activeCount, 1 To 3)
        For i = 1 To activeCount
            idx = activeIndices(i): p = p + 1
            out(p, 1) = batches(idx).Subtitle
            out(p, 2) = batches(idx).Counts(1)
            out(p, 3) = batches(idx).Counts(3)
        Next i
        ws.Range("A4").Resize(activeCount, 3).Value2 = out
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=area, XlListObjectHasHeaders:=xlYes)
        lo.Name = "GeneralReport": lo.TableStyle = "": lo.ShowTotals = False
        lo.ShowTableStyleRowStripes = False: lo.ShowTableStyleColumnStripes = False: lo.ShowAutoFilter = True
    Else
        ws.Range("A3:C3").Interior.Color = RGB(28, 55, 103)
        ws.Range("A3:C3").Font.Color = RGB(255, 255, 255): ws.Range("A3:C3").Font.Bold = True
    End If
    With ws.Range("A3:C3")
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .WrapText = True: .RowHeight = 36
    End With
    With ws.PageSetup
        .PrintArea = "$A$1:$C$" & lastRow: .Orientation = xlPortrait
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False: .CenterHorizontally = True
    End With
End Sub

Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir$(fullPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function

Private Function PathInUse(ByVal fullPath As String) As Boolean
    Dim book As Workbook
    If FileExists(fullPath) Then PathInUse = True: Exit Function
    For Each book In Application.Workbooks
        If StrComp(book.FullName, fullPath, vbTextCompare) = 0 Then PathInUse = True: Exit Function
    Next book
End Function

Private Function NewOutputPath(ByVal folder As String) As String
    Dim path As String, stem As String, n As Long
    stem = Left$(OUTPUT_FILE, Len(OUTPUT_FILE) - 5)
    path = folder & Application.PathSeparator & OUTPUT_FILE
    n = 1
    Do While PathInUse(path)
        n = n + 1
        path = folder & Application.PathSeparator & stem & "_" & Format$(n, "00") & ".xlsx"
    Loop
    NewOutputPath = path
End Function

Private Sub AddIssue(ByVal issues As Collection, ByRef total As Long, ByVal detail As String)
    total = total + 1
    If issues.Count < MAX_ISSUES Then issues.Add detail
End Sub

Private Function IssueSummary(ByVal issues As Collection, ByVal total As Long) As String
    Dim item As Variant
    IssueSummary = " (" & CStr(total) & ")"
    For Each item In issues
        IssueSummary = IssueSummary & vbLf & CStr(item)
    Next item
    If total > issues.Count Then IssueSummary = IssueSummary & vbLf & Text("more_issues") & " " & (total - issues.Count)
End Function

Private Function ColumnName(ByVal column As Long) As String
    Do While column > 0
        ColumnName = Chr$(65 + ((column - 1) Mod 26)) & ColumnName
        column = (column - 1) \ 26
    Loop
End Function

Private Function Normalize(ByVal value As Variant) As String
    Dim s As String, i As Long
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    s = LCase$(CStr(value))
    s = Replace(s, ChrW(&HA0), " ")
    s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " "): s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(&H623), ChrW(&H627))
    s = Replace(s, ChrW(&H625), ChrW(&H627))
    s = Replace(s, ChrW(&H622), ChrW(&H627))
    s = Replace(s, ChrW(&H649), ChrW(&H64A))
    s = Replace(s, ChrW(&H640), "")
    s = Replace(s, ChrW(&H200E), ""): s = Replace(s, ChrW(&H200F), "")
    For i = 0 To 9
        s = Replace(s, ChrW(&H660 + i), CStr(i))
        s = Replace(s, ChrW(&H6F0 + i), CStr(i))
    Next i
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Normalize = Trim$(s)
End Function

Private Function Txt(ByVal codes As String) As String
    Dim token As Variant
    For Each token In Split(codes, " ")
        If Len(CStr(token)) > 0 Then Txt = Txt & ChrW(CLng("&H" & CStr(token)))
    Next token
End Function

Private Function Text(ByVal key As String) As String
    Select Case key
    Case "save_tool": Text = Txt("627 62D 641 638 20 645 644 641 20 627 644 623 62F 648 627 62A 20 623 648 644 627 20 641 64A 20 627 644 645 62C 644 62F 20 627 644 630 64A 20 64A 62D 62A 648 64A 20 639 644 649 20 54 65 72 6D 30 32 52 65 73 75 6C 74 73 2E 78 6C 73 78 2E")
    Case "no_sheets": Text = Txt("644 627 20 62A 648 62C 62F 20 623 648 631 627 642 20 639 645 644 20 641 64A 20 627 644 645 644 641 20 627 644 645 635 62F 631 64A 2E")
    Case "bad_headers": Text = Txt("62A 639 630 631 20 62A 62D 62F 64A 62F 20 623 639 645 62F 629 20 627 644 645 639 62F 644 627 62A 20 641 64A 20 627 644 623 648 631 627 642 20 627 644 62A 627 644 64A 629 2E 20 644 645 20 64A 646 634 623 20 645 644 641 20 646 62A 627 626 62C 3A")
    Case "unique_headers": Text = Txt("64A 62C 628 20 648 62C 648 62F 20 639 645 648 62F 20 648 627 62D 62F 20 644 643 644 20 645 639 62F 644 20 641 64A 20 627 644 635 641 20 32")
    Case "term1": Text = Txt("645 639 62F 644 20 627 644 641 635 644 20 31")
    Case "term2": Text = Txt("645 639 62F 644 20 627 644 641 635 644 20 32")
    Case "difference": Text = Txt("627 644 641 627 631 642")
    Case "average_word": Text = Txt("645 639 62F 644")
    Case "term_word": Text = Txt("627 644 641 635 644")
    Case "first_word": Text = Txt("627 644 623 648 644")
    Case "second_word": Text = Txt("627 644 62B 627 646 64A")
    Case "working": Text = Txt("62A 62C 647 64A 632 20 627 644 648 631 642 629 3A")
    Case "failed": Text = Txt("62A 639 630 631 20 625 646 62C 627 632 20 627 644 645 642 627 631 646 629 3A")
    Case "done": Text = Txt("62A 645 20 625 646 634 627 621 20 645 644 641 20 627 644 645 642 627 631 646 629 3A")
    Case "sheets": Text = Txt("639 62F 62F 20 627 644 623 648 631 627 642")
    Case "pupils": Text = Txt("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 645 646 642 648 644 64A 646")
    Case "skipped": Text = Txt("62A 62C 627 648 632 20 627 644 628 631 646 627 645 62C 20 627 644 635 641 648 641 20 627 644 62A 627 644 64A 629 20 645 639 20 627 633 62A 645 631 627 631 20 646 642 644 20 627 644 635 641 648 641 20 627 644 633 644 64A 645 629 61B 20 627 644 623 639 645 62F 629 20 628 64A 646 20 627 644 642 648 633 64A 646 20 647 64A 20 645 648 627 636 639 20 627 644 62E 637 623 3A")
    Case "missing_source": Text = Txt("627 644 645 644 641 20 627 644 645 635 62F 631 64A 20 63A 64A 631 20 645 648 62C 648 62F 3A")
    Case "invalid_row": Text = Txt("627 633 645 20 63A 64A 631 20 635 627 644 62D 20 623 648 20 645 639 62F 644 20 641 627 631 63A 20 623 648 20 63A 64A 631 20 631 642 645 64A 20 623 648 20 62E 627 631 62C 20 627 644 645 62C 627 644 20 645 646 20 30 20 625 644 649 20 32 30")
    Case "title": Text = Txt("645 642 627 631 646 629 20 645 639 62F 644 64A 20 627 644 641 635 644 20 627 644 623 648 644 20 648 627 644 641 635 644 20 627 644 62B 627 646 64A")
    Case "group1": Text = Txt("62A 644 627 645 64A 630 20 645 62A 641 648 642 648 646")
    Case "low_list": Text = Txt("62A 644 627 645 64A 630 20 627 644 645 62A 631 627 62C 639 648 646")
    Case "general_report": Text = Txt("62A 642 631 64A 631 20 639 627 645")
    Case "pedagogical_group": Text = Txt("627 644 641 648 62C 20 627 644 62A 631 628 648 64A")
    Case "high_count": Text = Txt("639 62F 62F 20 627 644 645 62A 641 648 642 64A 646")
    Case "low_count": Text = Txt("639 62F 62F 20 627 644 645 62A 631 627 62C 639 64A 646")
    Case "group3": Text = Txt("62A 644 627 645 64A 630 20 645 62A 631 627 62C 639 648 646")
    Case "high_list": Text = Txt("642 627 626 645 629 20 627 644 645 62A 641 648 642 648 646")
    Case "class": Text = Txt("627 644 642 633 645")
    Case "too_many_rows": Text = Txt("639 62F 62F 20 627 644 635 641 648 641 20 64A 62A 62C 627 648 632 20 633 639 629 20 627 644 648 631 642 629 3A")
    Case "more_issues": Text = Txt("623 62E 637 627 621 20 625 636 627 641 64A 629 20 63A 64A 631 20 645 639 631 648 636 629 3A")
    End Select
End Function

' Stable sort of the first n records by the difference (column 8), largest first.
' Same difference = original order kept.
Private Sub SortRecordsDescending(ByRef values() As Variant, ByRef groups() As Long, ByRef sourceRows() As Long, ByVal n As Long)
    Dim i As Long, j As Long, c As Long, g As Long, s As Long, d As Double
    Dim rec(2 To 8) As Variant
    For i = 2 To n
        d = values(i, 8)
        For c = 2 To 8: rec(c) = values(i, c): Next c
        g = groups(i): s = sourceRows(i)
        j = i - 1
        Do While j >= 1
            If values(j, 8) >= d Then Exit Do
            For c = 2 To 8: values(j + 1, c) = values(j, c): Next c
            groups(j + 1) = groups(j): sourceRows(j + 1) = sourceRows(j)
            j = j - 1
        Loop
        For c = 2 To 8: values(j + 1, c) = rec(c): Next c
        groups(j + 1) = g: sourceRows(j + 1) = s
    Next i
End Sub

' order(1..n) = positions of keys(1..n), largest key first (stable).
Private Sub SortIndexDescending(ByRef keys() As Double, ByRef order() As Long, ByVal n As Long)
    Dim i As Long, j As Long, cur As Long
    For i = 1 To n: order(i) = i: Next i
    For i = 2 To n
        cur = order(i): j = i - 1
        Do While j >= 1
            If keys(order(j)) >= keys(cur) Then Exit Do
            order(j + 1) = order(j): j = j - 1
        Loop
        order(j + 1) = cur
    Next i
End Sub


' ==============================================================================
' OUTPUT FOLDER + FREE FILE NAME (nothing that already exists is replaced)
' ==============================================================================
' Name of the output folder created next to the tool workbook (hex code points)
Private Function OutputFolderName() As String
    OutputFolderName = AR("0645 0639 0627 0644 062C 0629 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0641 0635 0644 0020 0030 0032")
End Function

' ==============================================================================
' HELPERS
' ==============================================================================

' Builds a string from space-separated hex code points, e.g. AR("0627 0644")
Private Function AR(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    AR = s
End Function


' Removes invisible differences: NBSP, direction marks, tatweel, double spaces
Private Function NormalizeKey(ByVal s As String) As String
    Dim c As Variant
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    For Each c In Array(&H200B, &H200E, &H200F, &H202A, &H202B, &H202C, &H640)
        s = Replace(s, ChrW(CLng(c)), "")
    Next c
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeKey = Trim(s)
End Function

Private Function CellText(ByVal r As Range) As String
    If IsError(r.Value) Then CellText = "" Else CellText = CStr(r.Value)
End Function

Private Function FileExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FileExistsPath = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExistsPath = False: Err.Clear
End Function

Private Function FolderExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExistsPath = ((GetAttr(p) And vbDirectory) <> 0)
    If Err.Number <> 0 Then FolderExistsPath = False: Err.Clear
End Function

Private Function EnsureFolder(ByVal p As String) As Boolean
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    If Not FolderExistsPath(p) Then
        On Error Resume Next
        MkDir p
        Err.Clear
        On Error GoTo 0
    End If
    EnsureFolder = FolderExistsPath(p)
End Function

Private Function IsWorkbookOpen(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

' Last row containing anything, in any column (0 if the sheet is empty)
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If f Is Nothing Then LastUsedRow = 0 Else LastUsedRow = f.Row
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

' ==============================================================================
' COLLECTION OF THE SOURCE FILES INTO ONE WORKBOOK (fonts + dynamic tables)
' ==============================================================================
' Checks the Windows font registry (HKLM + HKCU). If the check itself cannot run,
' returns True (keeps the requested font instead of forcing the fallback).
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
                    If p > 0 Then nm = Left(nm, p - 1)          ' strip "(TrueType)" / "(OpenType)"
                    ' exact name, or family name of a variable font (e.g. "Bahnschrift" for "Bahnschrift SemiBold")
                    If StrComp(nm, FontName, vbTextCompare) = 0 Then FontExists = True: Exit Function
                    If StrComp(Left(FontName, Len(nm) + 1), nm & " ", vbTextCompare) = 0 Then FontExists = True: Exit Function
                Next i
            End If
        End If
    Next hv

    FontExists = Not enumOK     ' enumeration worked and nothing matched -> False; enumeration failed -> True
    Exit Function
Fallback:
    FontExists = True
End Function

Private Sub SetupFonts()
    If FontExists(FONT_ARABIC) Then g_ActiveArabicFont = FONT_ARABIC Else g_ActiveArabicFont = FONT_FALLBACK
    If FontExists(FONT_NUMERIC) Then g_ActiveNumericFont = FONT_NUMERIC Else g_ActiveNumericFont = FONT_FALLBACK
End Sub

' ==============================================================================
' STRUCTURE CHECK (engine). A results file has its identity headers in row
' DATA_START_ROW: B = name, C = birth date, D = sex. A file with any other layout
' (for example the wishes file, whose headers are in row 2) is skipped and named
' in the report; it never stops the run and it is never recognised by its file name.
' ==============================================================================
Private Function SrcHasResultsLayout(ByVal ws As Worksheet) As Boolean
    SrcHasResultsLayout = SrcHeaderHas(ws, 2, "627 633 645") _
                      And SrcHeaderHas(ws, 3, "645 64A 644 627 62F") _
                      And SrcHeaderHas(ws, 4, "62C 646 633")
End Function

Private Function SrcHeaderHas(ByVal ws As Worksheet, ByVal col As Long, ByVal wordCodes As String) As Boolean
    Dim v As Variant, h As String, w As String
    On Error Resume Next
    v = ws.Cells(DATA_START_ROW, col).MergeArea.Cells(1, 1).Value2
    On Error GoTo 0
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    h = SrcNormHeader(CStr(v))
    w = SrcNormHeader(SrcWord(wordCodes))
    SrcHeaderHas = (Len(w) > 0 And InStr(h, w) > 0)
End Function

Private Function SrcWord(ByVal codes As String) As String
    Dim p As Variant, s As String
    For Each p In Split(codes, " ")
        If Len(p) > 0 Then s = s & ChrW(CLng("&H" & p))
    Next p
    SrcWord = s
End Function

Private Function SrcNormHeader(ByVal t As String) As String
    t = Replace(t, ChrW(&H623), ChrW(&H627))    ' alef variants -> alef
    t = Replace(t, ChrW(&H625), ChrW(&H627))
    t = Replace(t, ChrW(&H622), ChrW(&H627))
    t = Replace(t, ChrW(&H649), ChrW(&H64A))    ' alef maksura -> yaa
    t = Replace(t, ChrW(&H629), ChrW(&H647))    ' taa marbuta -> haa
    t = Replace(t, ChrW(&H640), "")             ' tatweel
    t = Replace(t, ChrW(&HA0), " ")
    t = Replace(t, ChrW(&H200E), "")
    t = Replace(t, ChrW(&H200F), "")
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    SrcNormHeader = LCase$(Trim$(t))
End Function

' True when the workbook has the sheets of the wishes file (only used to word the report)
Private Function SrcHasWishSheets(ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, "Tecno", vbTextCompare) = 0 Or StrComp(ws.Name, "Litr", vbTextCompare) = 0 Then
            SrcHasWishSheets = True
            Exit Function
        End If
    Next ws
End Function

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

    ' ---- Structure check: not a results layout -> skipped and named in the report
    If Not SrcHasResultsLayout(wsSource) Then
        If SrcHasWishSheets(wbSource) Then
            ReasonOut = "WISHES FILE layout (sheets Tecno / Litr) - not a results file, ignored"
        Else
            ReasonOut = "DIFFERENT_STRUCTURE (identity headers not found in row " & DATA_START_ROW & ") - ignored"
        End If
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

    ' ---- Read source data
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

    lastDataRow = 1 + UBound(DataArr, 1)     ' header at row 2 + data rows

    ' ---- Number formats BEFORE writing data (so "@" columns store text)
    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        If LastCol >= 4 Then .Columns("D").NumberFormat = "@"
        If LastCol >= 5 Then .Columns("E").NumberFormat = "@"
        If LastCol >= 6 Then .Range(.Columns("F"), .Columns(LastCol)).NumberFormat = "0.00"
    End With

    wsResult.Range("A2").Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

    ' ---- Unique, non-empty headers (case-insensitive, like Excel tables)
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

    ' ---- Dynamic table (range ends exactly at the last data row)
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

    ' ---- Formatting
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

        ' Identity columns: AutoFit. Score columns: fixed width, rotated bold headers.
        idLastCol = IDENTITY_LAST_COL
        If LastCol < idLastCol Then idLastCol = LastCol
        .Range(.Columns(1), .Columns(idLastCol)).EntireColumn.AutoFit

        If LastCol >= 6 Then .Range(.Columns(IDENTITY_LAST_COL + 1), .Columns(LastCol)).ColumnWidth = COL_WIDTH

        If LastCol >= 4 Then
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Orientation = xlUpward
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Font.Bold = True
        End If
    End With


    ' Apply title formatting last so table column fonts do not override it.
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

    ' Collect names first (Dir cannot be nested with other Dir calls)
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
        SkippedList = "  - " & SOURCE_FOLDER_NAME & " folder is empty or contains no Excel files:" & vbCrLf & _
                      "    " & SourcePath & vbCrLf
        Exit Sub
    End If

    sorted = SortedNames(names)     ' deterministic sheet order (01, 02, ...)
    SetupFonts

    For i = LBound(sorted) To UBound(sorted)
        f = CStr(sorted(i))
        dotPos = InStrRev(f, ".")
        ext = LCase(Mid(f, dotPos + 1))
        baseName = Left(f, dotPos - 1)

        If StrComp(SourcePath & f, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            ' skip the macro workbook itself
        ElseIf IsOpenInExcel(SourcePath & f) Then
            ' never touch (or close) a file the user already has open
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

' ==============================================================================
' SHARED ENGINE: builds the collected workbook from \Source02\ IN MEMORY.
' Returns True with wbOut OPEN and UNSAVED (the caller must close it with
' SaveChanges:=False, or save it). Returns False when nothing usable was found;
' the reasons are in skippedList. The source files are only read, never changed.
' Calculation mode is not touched. AutomationSecurity / DisplayAlerts are restored.
' ==============================================================================
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
        skippedList = "  - " & SOURCE_FOLDER_NAME & " folder not found: " & srcFolder & vbCrLf
        Exit Function
    End If

    oldSec = Application.AutomationSecurity
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable   ' block Auto_Open in .xlsm/.xls
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

    ' 1. Capture and clean the text from A5
    rawText = CleanText(CStr(wsSource.Range("A5").Value))

    ' 2. Split at the end of the school year
    SplitAtSchoolYear rawText, partTitle, partGroup


    ' 3. Write and format both merged ranges
    WriteMergedCell wsTarget.Range("B1:F1"), partTitle
    WriteMergedCell wsTarget.Range("H1:Q1"), partGroup

    ' Merged cells do not auto-fit, so set the row height explicitly
    wsTarget.Rows(1).RowHeight = HEADER_ROW_HEIGHT

End Sub

'----------------------------------------------------------
' Cleans text: line breaks, non-breaking spaces, control
' characters and repeated spaces.
'----------------------------------------------------------
Private Function CleanText(ByVal s As String) As String

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(160), " ")                       ' non-breaking space
    s = Application.WorksheetFunction.Clean(s)          ' control characters
    CleanText = Application.WorksheetFunction.Trim(s)   ' extra spaces

End Function

'----------------------------------------------------------
' Splits text into:
'   part1 = everything up to the end of the school year
'   part2 = the remainder (class / group name)
' If no year is found, part1 = full text and part2 = "".
'----------------------------------------------------------
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

    ' Western (0-9) and Arabic-Indic (U+0660-U+0669) digits
    digit = "[0-9\u0660-\u0669]"
    ' Separators: slash, backslash, hyphen, en dash, em dash
    seps = "[/\\\-\u2013\u2014]"

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True

    ' First choice: full school-year pattern (e.g. 2025-2026)
    rx.Pattern = digit & "{4}\s*" & seps & "\s*" & digit & "{4}"

    If Not rx.Test(txt) Then
        ' Fallback: a single 4-digit year
        rx.Pattern = digit & "{4}"
        If Not rx.Test(txt) Then Exit Sub
    End If

    Set m = rx.Execute(txt)
    splitPos = m(0).FirstIndex + m(0).Length

    part1 = Trim(Left(txt, splitPos))

    If splitPos < Len(txt) Then
        part2 = Trim(Mid(txt, splitPos + 1))

        ' Remove any separator left at the start of the group name
        Do While Len(part2) > 0
            If InStr(1, "-:/" & ChrW(&H2013) & ChrW(&H2014), Left(part2, 1)) = 0 Then Exit Do
            part2 = Trim(Mid(part2, 2))
        Loop
    End If

End Sub

'----------------------------------------------------------
' Merges the range, writes the text and applies formatting.
'----------------------------------------------------------
Private Sub WriteMergedCell(ByVal rng As Range, ByVal txt As String)

    With rng
        .UnMerge                                ' remove any previous merge
        .ClearContents                          ' clear leftovers
        .Merge                                  ' merge the range
        .Value = txt
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL                   ' right-to-left reading order
        .Font.Name = GetAvailableFont(PREFERRED_FONT, FALLBACK_FONT)
        .Font.Size = HEADER_FONT_SIZE
        .Font.Bold = True
    End With

End Sub

'----------------------------------------------------------
' Returns the preferred font if installed, otherwise the fallback.
'----------------------------------------------------------
Private Function GetAvailableFont(ByVal wanted As String, _
                                  ByVal fallback As String) As String

    If IsFontInstalled(wanted) Then
        GetAvailableFont = wanted
    Else
        GetAvailableFont = fallback
    End If

End Function

'----------------------------------------------------------
' Checks the system and per-user font folders for a file whose
' name starts with the font name (e.g. Cairo-Regular.ttf).
'----------------------------------------------------------
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
