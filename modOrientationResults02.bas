Attribute VB_Name = "modOrientationResults02"
Option Explicit

' ============================================================================
' Module: modOrientationResults02
' Purpose:
'   Export the final guidance results from Term02Guidance.xlsx to a separate
'   workbook inside folder "تحليل نتائج الفصل الثاني" with sequential naming:
'   OrientationResults02.xlsx, OrientationResults02_01.xlsx, OrientationResults02_02.xlsx, ...
'
' Output:
'   - Saved inside subfolder "تحليل نتائج الفصل الثاني".
'   - Never overwrites existing files; creates sequential _01, _02 names.
'   - Sheet "Litr": all arts students from Data01.
'   - One sheet per science section from Data, named Tecno1, Tecno2, ...
'   - AutoFits all columns automatically.
'   - Protects all output sheets.
'
' Entry point: Run_OrientationResults02
' ============================================================================

Private m_runSucceeded As Boolean
Private m_silent As Boolean
Private m_lastReport As String
Private m_lastOutputPath As String

Private Const SOURCE_FILE As String = "Term02Guidance.xlsx"
Private Const OUTPUT_BASE_NAME As String = "OrientationResults02"
Private Const OUTPUT_EXT As String = ".xlsx"

' Folder name: "تحليل نتائج الفصل الثاني" stored as hex code points
Private Const OUTPUT_FOLDER_CODES As String = "62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A"

Private Const SCIENCE_SHEET As String = "Data"
Private Const ARTS_SHEET As String = "Data01"

Private Const FIRST_DATA_ROW As Long = 3
Private Const HEADER_ROW As Long = 2

Private Const SCIENCE_WIDTH As Long = 21         ' source table width (A:U)
Private Const SCIENCE_OUT_WIDTH As Long = 20     ' exported width: section column removed
Private Const ARTS_WIDTH As Long = 15

Private Const SCIENCE_SECTION_COL As Long = 14   ' N
Private Const SCIENCE_AVERAGE_COL As Long = 13   ' M

Private Const ARTS_OUTPUT_SHEET As String = "Litr"
Private Const SCIENCE_SHEET_PREFIX As String = "Tecno"

' Row-1 titles as hex code points. Change TITLE_TERM_WORD to switch the term word:
'   second term = "627 644 62B 627 646 64A"   first term = "627 644 623 648 644"
Private Const TITLE_TERM_WORD As String = "627 644 62B 627 646 64A"
Private Const TITLE_ARTS_PRE As String = "627 644 62A 648 62C 64A 647 20 627 644 62A 62F 631 64A 62C 64A 20 644 644 641 635 644 20"
Private Const TITLE_ARTS_POST As String = "20 2D 20 623 648 644 649 20 62B 627 646 648 64A 20 62C 630 639 20 645 634 62A 631 643 20 622 62F 627 628"
Private Const TITLE_SCI_PRE As String = "645 62D 636 631 20 627 644 62A 648 62C 64A 647 20 627 644 62A 62F 631 64A 62C 64A 20 62D 633 628 20 646 62A 627 626 62C 20 627 644 641 635 644 20"
Private Const TITLE_SCI_POST As String = "20 644 644 633 646 629 20 627 644 623 648 644 649 20 62C 630 639 20 645 634 62A 631 643 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627 20"

Private Const FONT_ARABIC As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK As String = "Calibri"
Private Const TITLE_FONT As String = "Cairo"
Private Const TITLE_FONT_FALLBACK As String = "Arial"
Private Const BODY_FONT_SIZE As Long = 14
Private Const TITLE_FONT_SIZE As Long = 11

Public Sub SetSilentMode(ByVal silent As Boolean)
    m_silent = silent
End Sub

Public Function LastReport() As String
    LastReport = m_lastReport
End Function

Public Function OrientationResults02Succeeded() As Boolean
    OrientationResults02Succeeded = m_runSucceeded
End Function

Public Function Results02OutputPath() As String
    Results02OutputPath = m_lastOutputPath
End Function

' ---------------------------------------------------------------------------
' Public entry point
' ---------------------------------------------------------------------------
Public Sub Run_OrientationResults02()

    Dim wbSource As Workbook
    Dim wbOutput As Workbook
    Dim wsScience As Worksheet
    Dim wsArts As Worksheet
    Dim loScience As ListObject
    Dim loArts As ListObject

    Dim sourceOpened As Boolean
    Dim stateSaved As Boolean
    Dim wasSilent As Boolean

    Dim targetFolder As String
    Dim outputPath As String
    Dim artsCount As Long
    Dim scienceCount As Long
    Dim sectionCount As Long
    Dim skippedScience As Long
    Dim reportText As String
    Dim failed As Boolean
    Dim errText As String

    wasSilent = m_silent
    m_runSucceeded = False
    m_lastReport = ""
    m_lastOutputPath = ""

    On Error GoTo EH
    Application.EnableCancelKey = xlErrorHandler

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 1000, , _
                  "Save the tool workbook before running Orientation Results."
    End If

    targetFolder = EnsureOutputFolder()
    outputPath = GetNextAvailableOutputPath(targetFolder)

    stateSaved = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wbSource = OpenSourceWorkbook(SOURCE_FILE, sourceOpened)

    Set wsScience = FindSheet(wbSource, SCIENCE_SHEET)
    Set wsArts = FindSheet(wbSource, ARTS_SHEET)

    If wsScience Is Nothing And wsArts Is Nothing Then
        Err.Raise vbObjectError + 1002, , _
                  "Neither Data nor Data01 was found in " & SOURCE_FILE & "."
    End If

    If Not wsScience Is Nothing Then
        Set loScience = FindGuidanceTable(wsScience, SCIENCE_WIDTH)
    End If

    If Not wsArts Is Nothing Then
        Set loArts = FindGuidanceTable(wsArts, ARTS_WIDTH)
    End If

    Set wbOutput = Workbooks.Add(xlWBATWorksheet)

    ' Arts are written first to the sheet named "Litr".
    If Not wsArts Is Nothing Then
        artsCount = WriteArtsResults(wbSource, wsArts, loArts, wbOutput)
    End If

    ' Science is split into one sheet per section.
    If Not wsScience Is Nothing Then
        WriteScienceResults wbSource, wsScience, loScience, wbOutput, _
                            scienceCount, sectionCount, skippedScience
    End If

    If wbOutput.Worksheets.Count = 1 Then
        If wbOutput.Worksheets(1).UsedRange.Cells.Count = 1 _
           And Len(CStr(wbOutput.Worksheets(1).Range("A1").Value2)) = 0 Then
            If artsCount = 0 And sectionCount = 0 Then
                Err.Raise vbObjectError + 1003, , _
                          "No student rows were found in Data or Data01."
            End If
        End If
    End If

    wbOutput.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing

    m_lastOutputPath = outputPath
    m_runSucceeded = True

    reportText = "Arts students (Litr): " & CStr(artsCount) & _
                 " | Science students: " & CStr(scienceCount) & _
                 " (" & CStr(sectionCount) & " section sheets)"

    If skippedScience > 0 Then
        reportText = reportText & vbCrLf & _
                     "Science rows without a valid section: " & CStr(skippedScience)
    End If

    reportText = reportText & vbCrLf & "Saved to: " & outputPath
    m_lastReport = reportText

CleanExit:
    On Error Resume Next

    If Not wbOutput Is Nothing Then
        wbOutput.Close SaveChanges:=False
    End If

    If sourceOpened And Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If

    If stateSaved Then
        Application.Calculation = xlCalculationAutomatic
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
            MsgBox "Orientation Results failed:" & vbCrLf & errText, _
                   vbCritical, "Orientation Results 02"
        End If
    Else
        If Not wasSilent Then
            MsgBox reportText, vbInformation, "Orientation Results 02"
        End If
    End If

    Exit Sub

EH:
    failed = True
    If Err.Number = 18 Then
        errText = "Execution was canceled by the user."
    Else
        errText = Err.Description
    End If
    Resume CleanExit

End Sub

' ---------------------------------------------------------------------------
' Output folder and sequential file naming (_01, _02, ...)
' ---------------------------------------------------------------------------
Private Function OutputFolderName() As String
    OutputFolderName = ArText(OUTPUT_FOLDER_CODES)
End Function

Private Function EnsureOutputFolder() As String
    Dim fso As Object
    Dim folderPath As String

    folderPath = ThisWorkbook.Path
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
    folderPath = folderPath & OutputFolderName()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    EnsureOutputFolder = folderPath
End Function

Private Function GetNextAvailableOutputPath(ByVal folderPath As String) As String
    Dim fso As Object
    Dim candidateName As String
    Dim candidatePath As String
    Dim seq As Long
    Dim sep As String

    sep = Application.PathSeparator
    If Right$(folderPath, 1) <> sep Then folderPath = folderPath & sep

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' First try OrientationResults02.xlsx; if it exists or is open, try _01, _02, ...
    candidateName = OUTPUT_BASE_NAME & OUTPUT_EXT
    candidatePath = folderPath & candidateName
    If Not fso.FileExists(candidatePath) And Not IsWorkbookNameOpen(candidateName) Then
        GetNextAvailableOutputPath = candidatePath
        Exit Function
    End If

    seq = 1
    Do
        candidateName = OUTPUT_BASE_NAME & "_" & Format$(seq, "00") & OUTPUT_EXT
        candidatePath = folderPath & candidateName
        If Not fso.FileExists(candidatePath) And Not IsWorkbookNameOpen(candidateName) Then
            GetNextAvailableOutputPath = candidatePath
            Exit Function
        End If
        seq = seq + 1
    Loop
End Function

Private Function IsWorkbookNameOpen(ByVal fileName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            IsWorkbookNameOpen = True
            Exit Function
        End If
    Next wb
End Function

' ---------------------------------------------------------------------------
' Arts output: one sheet named "Litr", all arts students, original source order.
' ---------------------------------------------------------------------------
Private Function WriteArtsResults(ByVal wbSource As Workbook, _
                                  ByVal wsSource As Worksheet, _
                                  ByVal loSource As ListObject, _
                                  ByVal wbOutput As Workbook) As Long

    Dim wsOutput As Worksheet
    Dim sourceData As Variant
    Dim headers As Variant
    Dim rowCount As Long

    headers = loSource.HeaderRowRange.Value2

    If loSource.DataBodyRange Is Nothing Then
        rowCount = 0
    Else
        sourceData = loSource.DataBodyRange.Value2
        rowCount = UBound(sourceData, 1)
    End If

    Set wsOutput = GetAvailableOutputSheet(wbOutput, ARTS_OUTPUT_SHEET)
    wsOutput.Name = ARTS_OUTPUT_SHEET

    WriteOutputSheet wsSource, wsOutput, headers, sourceData, _
                     rowCount, ARTS_WIDTH, ArtsTitle(), False

    WriteArtsResults = rowCount

End Function

' ---------------------------------------------------------------------------
' Science output: one worksheet per section (Tecno1, Tecno2, ...).
' ---------------------------------------------------------------------------
Private Sub WriteScienceResults(ByVal wbSource As Workbook, _
                                ByVal wsSource As Worksheet, _
                                ByVal loSource As ListObject, _
                                ByVal wbOutput As Workbook, _
                                ByRef totalStudents As Long, _
                                ByRef sectionSheets As Long, _
                                ByRef skippedRows As Long)

    Dim sourceData As Variant
    Dim headers As Variant
    Dim groups As Object
    Dim rowList As Collection
    Dim keys As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim sectionNo As Long
    Dim key As String
    Dim wsOutput As Worksheet
    Dim groupIndex As Long
    Dim sectionName As String

    If loSource.DataBodyRange Is Nothing Then Exit Sub

    sourceData = loSource.DataBodyRange.Value2
    rowCount = UBound(sourceData, 1)
    headers = loSource.HeaderRowRange.Value2

    Set groups = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        If TryGetSectionNumber(sourceData(i, SCIENCE_SECTION_COL), sectionNo) Then
            key = CStr(sectionNo)

            If Not groups.Exists(key) Then
                Set rowList = New Collection
                groups.Add key, rowList
            End If

            Set rowList = groups(key)
            rowList.Add i
        Else
            skippedRows = skippedRows + 1
        End If
    Next i

    If groups.Count = 0 Then Exit Sub

    keys = groups.Keys
    SortNumericKeys keys

    For groupIndex = LBound(keys) To UBound(keys)

        sectionNo = CLng(keys(groupIndex))
        sectionName = SCIENCE_SHEET_PREFIX & CStr(sectionNo)

        Set rowList = groups(CStr(sectionNo))
        Set wsOutput = GetAvailableOutputSheet(wbOutput, sectionName)

        WriteScienceGroup wsSource, wsOutput, headers, sourceData, _
                          rowList, sectionNo

        totalStudents = totalStudents + rowList.Count
        sectionSheets = sectionSheets + 1

    Next groupIndex

End Sub

' ---------------------------------------------------------------------------
' Build one science section and sort by term average descending.
' ---------------------------------------------------------------------------
Private Sub WriteScienceGroup(ByVal wsSource As Worksheet, _
                              ByVal wsOutput As Worksheet, _
                              ByVal headers As Variant, _
                              ByVal sourceData As Variant, _
                              ByVal rowList As Collection, _
                              ByVal sectionNo As Long)

    Dim order() As Long
    Dim outputData() As Variant
    Dim outputHeaders() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim c As Long
    Dim sourceRow As Long

    rowCount = rowList.Count
    ReDim order(1 To rowCount)

    For i = 1 To rowCount
        order(i) = CLng(rowList(i))
    Next i

    StableSortRowsByAverage order, sourceData

    ReDim outputHeaders(1 To 1, 1 To SCIENCE_OUT_WIDTH)
    For c = 1 To SCIENCE_OUT_WIDTH
        outputHeaders(1, c) = headers(1, SourceColumn(c, True))
    Next c

    ReDim outputData(1 To rowCount, 1 To SCIENCE_OUT_WIDTH)

    For i = 1 To rowCount
        sourceRow = order(i)

        For c = 1 To SCIENCE_OUT_WIDTH
            outputData(i, c) = sourceData(sourceRow, SourceColumn(c, True))
        Next c

        ' The first column is the rank within this section.
        outputData(i, 1) = i
    Next i

    WriteOutputSheet wsSource, wsOutput, outputHeaders, outputData, _
                     rowCount, SCIENCE_OUT_WIDTH, ScienceTitle(sectionNo), True

End Sub

Private Function SourceColumn(ByVal outCol As Long, ByVal isScience As Boolean) As Long
    If isScience And outCol >= SCIENCE_SECTION_COL Then
        SourceColumn = outCol + 1
    Else
        SourceColumn = outCol
    End If
End Function

' ---------------------------------------------------------------------------
' Titles (row 1)
' ---------------------------------------------------------------------------
Private Function ArText(ByVal codes As String) As String
    Dim parts() As String
    Dim i As Long
    Dim s As String

    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    ArText = s
End Function

Private Function ArtsTitle() As String
    ArtsTitle = ArText(TITLE_ARTS_PRE) & ArText(TITLE_TERM_WORD) & ArText(TITLE_ARTS_POST)
End Function

Private Function ScienceTitle(ByVal sectionNo As Long) As String
    ScienceTitle = ArText(TITLE_SCI_PRE) & ArText(TITLE_TERM_WORD) & _
                   ArText(TITLE_SCI_POST) & Format$(sectionNo, "00")
End Function

' ---------------------------------------------------------------------------
' Write data, AutoFit all columns automatically, format title row, and protect
' the output worksheet.
' ---------------------------------------------------------------------------
Private Sub WriteOutputSheet(ByVal wsSource As Worksheet, _
                             ByVal wsOutput As Worksheet, _
                             ByVal headers As Variant, _
                             ByVal data As Variant, _
                             ByVal rowCount As Long, _
                             ByVal columnCount As Long, _
                             ByVal titleText As String, _
                             ByVal isScience As Boolean)

    Dim lastRow As Long
    Dim tbl As ListObject
    Dim targetRange As Range
    Dim targetTitle As Range
    Dim c As Long
    Dim srcCol As Long
    Dim arabicFont As String
    Dim numericFont As String
    Dim titleFont As String
    Dim wishStartCol As Long
    Dim decisionCol As Long

    wsOutput.Unprotect Password:=""
    wsOutput.Cells.Clear
    wsOutput.DisplayRightToLeft = True

    lastRow = HEADER_ROW + rowCount

    ' Header values come from the guidance table (section column already removed).
    wsOutput.Cells(HEADER_ROW, 1).Resize(1, columnCount).Value2 = headers

    For c = 1 To columnCount
        srcCol = SourceColumn(c, isScience)
        wsSource.Cells(HEADER_ROW, srcCol).Copy
        wsOutput.Cells(HEADER_ROW, c).PasteSpecial Paste:=xlPasteFormats
    Next c
    Application.CutCopyMode = False

    ' Data rows are copied as values only.
    If rowCount > 0 Then
        wsOutput.Cells(FIRST_DATA_ROW, 1).Resize(rowCount, columnCount).Value2 = data

        For c = 1 To columnCount
            srcCol = SourceColumn(c, isScience)
            wsSource.Cells(FIRST_DATA_ROW, srcCol).Copy
            wsOutput.Range(wsOutput.Cells(FIRST_DATA_ROW, c), _
                           wsOutput.Cells(lastRow, c)).PasteSpecial Paste:=xlPasteFormats
        Next c
        Application.CutCopyMode = False

        ' Keep the source number formats for every output column.
        For c = 1 To columnCount
            srcCol = SourceColumn(c, isScience)
            wsOutput.Columns(c).NumberFormat = _
                wsSource.Cells(FIRST_DATA_ROW, srcCol).NumberFormat
        Next c
    End If

    If isScience And rowCount > 0 Then
        wsOutput.Columns(1).NumberFormat = "0"
    End If

    Set targetRange = wsOutput.Range(wsOutput.Cells(HEADER_ROW, 1), _
                                     wsOutput.Cells(lastRow, columnCount))

    Set tbl = wsOutput.ListObjects.Add(SourceType:=xlSrcRange, _
                                       Source:=targetRange, _
                                       XlListObjectHasHeaders:=xlYes)

    On Error Resume Next
    tbl.TableStyle = ""
    tbl.ShowTableStyleColumnStripes = False
    tbl.ShowTableStyleRowStripes = False
    tbl.ShowTableStyleFirstColumn = False
    tbl.ShowTableStyleLastColumn = False
    tbl.ShowAutoFilter = True
    On Error GoTo 0

    ' Collector-style typography and alignment.
    arabicFont = AvailableFont(FONT_ARABIC, FONT_FALLBACK)
    numericFont = AvailableFont(FONT_NUMERIC, FONT_FALLBACK)

    If isScience Then
        wishStartCol = 14
        decisionCol = 20
    Else
        wishStartCol = 11
        decisionCol = 15
    End If

    With wsOutput
        .Range(.Cells(HEADER_ROW, 1), .Cells(lastRow, columnCount)).HorizontalAlignment = xlCenter
        .Range(.Cells(HEADER_ROW, 1), .Cells(lastRow, columnCount)).VerticalAlignment = xlCenter
        .Range(.Cells(HEADER_ROW, 1), .Cells(lastRow, columnCount)).WrapText = False

        .Columns(1).Font.Name = numericFont
        .Columns(1).Font.Size = BODY_FONT_SIZE

        .Columns("B").Font.Name = arabicFont
        .Columns("B").Font.Size = BODY_FONT_SIZE

        .Columns("C").Font.Name = numericFont
        .Columns("C").Font.Size = BODY_FONT_SIZE

        If columnCount >= 4 Then
            .Range(.Columns(4), .Columns(5)).Font.Name = arabicFont
            .Range(.Columns(4), .Columns(5)).Font.Size = BODY_FONT_SIZE
        End If

        If columnCount >= 6 Then
            .Range(.Columns(6), .Columns(columnCount)).Font.Name = numericFont
            .Range(.Columns(6), .Columns(columnCount)).Font.Size = BODY_FONT_SIZE
        End If

        ' Ensure Arabic text columns (wishes and decision) use the Arabic font
        If columnCount >= wishStartCol + 1 Then
            .Range(.Columns(wishStartCol), .Columns(wishStartCol + 1)).Font.Name = arabicFont
            .Range(.Columns(wishStartCol), .Columns(wishStartCol + 1)).Font.Size = BODY_FONT_SIZE
        End If
        If columnCount >= decisionCol Then
            .Columns(decisionCol).Font.Name = arabicFont
            .Columns(decisionCol).Font.Size = BODY_FONT_SIZE
        End If

        If columnCount >= 4 Then
            .Range(.Cells(HEADER_ROW, 4), .Cells(HEADER_ROW, columnCount)).Orientation = xlUpward
            .Range(.Cells(HEADER_ROW, 4), .Cells(HEADER_ROW, columnCount)).Font.Bold = True
        End If

        .Rows(HEADER_ROW).RowHeight = 120
        If rowCount > 0 Then
            .Range(.Rows(FIRST_DATA_ROW), .Rows(lastRow)).RowHeight = 20
        End If

        ' AutoFit all columns automatically before merging Row 1
        .Range(.Columns(1), .Columns(columnCount)).EntireColumn.AutoFit
    End With

    ' Title row: merged and formatted after AutoFit so it does not distort column widths.
    Set targetTitle = wsOutput.Range(wsOutput.Cells(1, 1), _
                                     wsOutput.Cells(1, columnCount))
    targetTitle.Merge
    wsOutput.Cells(1, 1).Value2 = titleText

    On Error Resume Next
    With wsSource.Cells(1, 1)
        If .Interior.Pattern <> xlNone Then
            targetTitle.Interior.Pattern = .Interior.Pattern
            targetTitle.Interior.Color = .Interior.Color
        End If
        targetTitle.Font.Color = .Font.Color
    End With
    On Error GoTo 0

    titleFont = AvailableFont(TITLE_FONT, TITLE_FONT_FALLBACK)
    With targetTitle
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Font.Name = titleFont
        .Font.Size = TITLE_FONT_SIZE
        .Font.Bold = True
    End With
    wsOutput.Rows(1).RowHeight = 30

    ' Protect every output sheet
    wsOutput.Cells.Locked = True
    wsOutput.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

End Sub

' ---------------------------------------------------------------------------
' Workbook and table helpers
' ---------------------------------------------------------------------------
Private Function OpenSourceWorkbook(ByVal fileName As String, _
                                    ByRef openedHere As Boolean) As Workbook

    Dim wb As Workbook
    Dim fullPath As String

    openedHere = False
    fullPath = ThisWorkbook.Path & Application.PathSeparator & fileName

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenSourceWorkbook = wb
            Exit Function
        End If

        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1010, , _
                      "A different workbook with the same name is already open: " & wb.FullName
        End If
    Next wb

    If Len(Dir$(fullPath)) = 0 Then
        Err.Raise vbObjectError + 1011, , "Source file not found: " & fullPath
    End If

    Set OpenSourceWorkbook = Workbooks.Open( _
        Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    openedHere = True

End Function

Private Function FindSheet(ByVal wb As Workbook, _
                           ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FindSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindGuidanceTable(ByVal ws As Worksheet, _
                                   ByVal requiredWidth As Long) As ListObject

    Dim lo As ListObject

    For Each lo In ws.ListObjects
        If lo.Range.Row = HEADER_ROW And lo.Range.Column = 1 Then
            If lo.ListColumns.Count < requiredWidth Then
                Err.Raise vbObjectError + 1020, , _
                          "Table on sheet '" & ws.Name & "' has fewer than " & _
                          CStr(requiredWidth) & " columns."
            End If

            Set FindGuidanceTable = lo
            Exit Function
        End If
    Next lo

    Err.Raise vbObjectError + 1021, , _
              "A guidance table starting at A2 was not found on sheet '" & ws.Name & "'."

End Function

Private Function GetAvailableOutputSheet(ByVal wb As Workbook, _
                                         ByVal desiredName As String) As Worksheet

    Dim ws As Worksheet

    If wb.Worksheets.Count = 1 Then
        Set ws = wb.Worksheets(1)
        If ws.UsedRange.Cells.Count = 1 And Len(CStr(ws.Range("A1").Value2)) = 0 _
           And ws.ListObjects.Count = 0 And Not ws.ProtectContents Then
            ws.Name = desiredName
            Set GetAvailableOutputSheet = ws
            Exit Function
        End If
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = desiredName
    Set GetAvailableOutputSheet = ws

End Function

' ---------------------------------------------------------------------------
' Section grouping and stable sorting
' ---------------------------------------------------------------------------
Private Function TryGetSectionNumber(ByVal value As Variant, _
                                     ByRef sectionNo As Long) As Boolean

    Dim s As String

    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function

    s = Trim$(CStr(value))
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    If CDbl(s) < 1 Or CDbl(s) <> Fix(CDbl(s)) Then Exit Function

    sectionNo = CLng(s)
    TryGetSectionNumber = True

End Function

Private Sub SortNumericKeys(ByRef keys As Variant)

    Dim i As Long
    Dim j As Long
    Dim temp As Variant

    If Not IsArray(keys) Then Exit Sub

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(j)) < CLng(keys(i)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i

End Sub

Private Sub StableSortRowsByAverage(ByRef order() As Long, _
                                    ByVal sourceData As Variant)

    Dim i As Long
    Dim j As Long
    Dim currentRow As Long
    Dim currentAverage As Double

    For i = LBound(order) + 1 To UBound(order)

        currentRow = order(i)
        currentAverage = NumericValue(sourceData(currentRow, SCIENCE_AVERAGE_COL))
        j = i - 1

        Do While j >= LBound(order)
            If NumericValue(sourceData(order(j), SCIENCE_AVERAGE_COL)) >= currentAverage Then
                Exit Do
            End If

            order(j + 1) = order(j)
            j = j - 1
        Loop

        order(j + 1) = currentRow

    Next i

End Sub

Private Function NumericValue(ByVal value As Variant) As Double

    If IsError(value) Or IsEmpty(value) Then
        NumericValue = -1E+99
    ElseIf IsNumeric(value) Then
        NumericValue = CDbl(value)
    Else
        NumericValue = -1E+99
    End If

End Function

Private Function AvailableFont(ByVal preferred As String, _
                               ByVal fallback As String) As String
    AvailableFont = preferred
End Function