Attribute VB_Name = "modBiuldAnalysisfilesCommon02"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesCommon02
' Purpose : Shared constants, text helpers, fonts, cell styles and the
'           "discover what is inside the Data sheet" routine used by
'           modBiuldAnalysisfilesAnalysis02 and modBiuldAnalysisfilesCartTec02.
' Notes   : - Pure ASCII: every Arabic text is built from hex code points (baf_ArT).
'           - Formats / fonts / colours come from modExamenAnalysis02.
'==============================================================================

' ---- sheet / table names
Public Const baf_DATA_SHEET As String = "Data"
Public Const baf_DATA_TABLE As String = "Data"
Public Const baf_DATA_HDR_ROW As Long = 6
Public Const baf_DATA_FIRST_ROW As Long = 7
Public Const baf_ANALYSIS_PREFIX As String = "Analysis"
Public Const baf_CART_SHEET As String = "Cart_Tec"
Public Const baf_COMPARE_SHEET As String = "Compare"

' ---- number formats (same as modExamenAnalysis02)
Public Const baf_FMT_COUNT As String = "00"
Public Const baf_FMT_DEC As String = "00.00"
Public Const baf_FMT_PCT As String = "[=1]%0;%00.00"
Public Const baf_PASS_MARK As Double = 10

' ---- layout (same as modExamenAnalysis02)
Public Const baf_FIRST_NUM_COL As Long = 4
Public Const baf_TITLE_ROW_HEIGHT As Double = 40
Public Const baf_DATA_ROW_HEIGHT As Double = 24
Public Const baf_HDR_ROW_HEIGHT As Double = 100
Public Const baf_ID_COL_WIDTH As Double = 4.3
Public Const baf_NUM_COL_WIDTH As Double = 6.3
Public Const baf_DATE_COL_MIN_WIDTH As Double = 10.5

' ---- fonts (same as modExamenAnalysis02)
Public Const baf_CHECK_FONTS As Boolean = True
Public Const baf_FONT_BASE As String = "Dubai"
Public Const baf_FONT_NAME As String = "Dubai Medium"
Public Const baf_FONT_NUM As String = "Aptos Slab"
Public Const baf_FONT_SAFE_TEXT As String = "Arial"
Public Const baf_FONT_SAFE_NUM As String = "Calibri"

Public Const baf_QT As String = """"

Public baf_gFontBase As String
Public baf_gFontName As String
Public baf_gFontNum As String
Public baf_gFontNotes As String

' ---- result of baf_DiscoverData (read the Data table once, share with the builders)
Public baf_gHdr() As String          ' table headers, 1..gNCols
Public baf_gNCols As Long
Public baf_gRows As Long
Public baf_gSecCol As Long           ' column of the section header
Public baf_gSubjN As Long            ' number of analysed subjects
Public baf_gSubjName() As String     ' subject name without the term suffix
Public baf_gSubjC1() As Long         ' column of term 1 mark
Public baf_gSubjC2() As Long         ' column of term 2 mark
Public baf_gAvgC1 As Long            ' column of term 1 average
Public baf_gAvgC2 As Long            ' column of term 2 average
Public baf_gPEIdx As Long            ' index of the physical-education subject (0 = none)
Public baf_gSecN As Long             ' number of sections that have students
Public baf_gSecList() As Long        ' their numbers
Public baf_gClassTitle As String     ' class title (Data!A5)

' ---- layout of the Analysis blocks (filled by baf_BuildAnalysisSheets, read by the Compare sheet)
Public baf_gBlkN As Long
Public baf_gBlkSec() As Long         ' section number of each block
Public baf_gBlkR0() As Long          ' first row (banner) of each block
Public baf_gBlkH1() As Long          ' header row of table 1; the 7 rows below = average ... lowest
Public baf_gBlkQ1() As Long          ' first row of the "passed" cells
Public baf_gBlkNQ() As Long          ' number of "passed" cells of each block
Public baf_gBlkQLbl() As String      ' their labels  (block, 1..4)

'==============================================================================
' TEXT HELPERS
'==============================================================================
Public Function baf_ArT(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    baf_ArT = s
End Function

Public Function baf_NormKey(ByVal s As String) As String
    Dim c As Variant
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    For Each c In Array(&H200B, &H200E, &H200F, &H202A, &H202B, &H202C, &H640)
        s = Replace(s, ChrW(CLng(c)), "")
    Next c
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    baf_NormKey = Trim$(s)
End Function

' Normalised Arabic used only for keyword matching (alef/ya forms)
Public Function baf_NormAr(ByVal s As String) As String
    s = baf_NormKey(s)
    s = Replace(s, ChrW(&H623), ChrW(&H627))
    s = Replace(s, ChrW(&H625), ChrW(&H627))
    s = Replace(s, ChrW(&H622), ChrW(&H627))
    s = Replace(s, ChrW(&H671), ChrW(&H627))
    s = Replace(s, ChrW(&H649), ChrW(&H64A))
    baf_NormAr = s
End Function

' Structured reference to one column of the Data table (always double brackets)
Public Function baf_ColRef(ByVal hdr As String) As String
    Dim t As String
    t = hdr
    t = Replace(t, "'", "''")
    t = Replace(t, "[", "'[")
    t = Replace(t, "]", "']")
    t = Replace(t, "#", "'#")
    baf_ColRef = baf_DATA_TABLE & "[[" & t & "]]"
End Function

Public Function baf_ColLetter(ByVal c As Long) As String
    Dim a As String
    a = Cells(1, c).Address(False, False)
    baf_ColLetter = Left$(a, Len(a) - 1)
End Function

' Replaces {key} by value; arguments come in pairs: key1, value1, key2, value2 ...
Public Function baf_Subst(ByVal f As String, ParamArray kv() As Variant) As String
    Dim i As Long
    For i = LBound(kv) To UBound(kv) - 1 Step 2
        f = Replace(f, "{" & CStr(kv(i)) & "}", CStr(kv(i + 1)))
    Next i
    baf_Subst = f
End Function

'==============================================================================
' FILES (FileSystemObject: safe with Arabic names)
'==============================================================================
Public Function baf_FolderExistsP(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    baf_FolderExistsP = fso.FolderExists(p)
    If Err.Number <> 0 Then baf_FolderExistsP = False: Err.Clear
End Function

Public Function baf_FileExistsP(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    baf_FileExistsP = fso.FileExists(p)
    If Err.Number <> 0 Then baf_FileExistsP = False: Err.Clear
End Function

Public Function baf_EnsureFolderP(ByVal p As String) As Boolean
    Dim fso As Object
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Not baf_FolderExistsP(p) Then
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder p
        Err.Clear
        On Error GoTo 0
    End If
    baf_EnsureFolderP = baf_FolderExistsP(p)
End Function

Private Function IsOpenWb(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsOpenWb = True
            Exit Function
        End If
    Next wb
End Function

' stem.xlsx, and when it already exists: stem_01.xlsx, stem_02.xlsx ... (never overwrites)
Public Function baf_NextFreePath(ByVal folder As String, ByVal stem As String, ByVal ext As String) As String
    Dim p As String, n As Long
    p = folder & stem & ext
    Do While baf_FileExistsP(p) Or IsOpenWb(p)
        n = n + 1
        p = folder & stem & "_" & Format$(n, "00") & ext
    Loop
    baf_NextFreePath = p
End Function

'==============================================================================
' FONTS
'==============================================================================
Public Sub baf_ResolveFonts(ByVal ws As Worksheet)
    baf_gFontNotes = ""
    baf_gFontBase = PickFont(ws, baf_FONT_BASE, baf_FONT_SAFE_TEXT)
    baf_gFontName = PickFont(ws, baf_FONT_NAME, baf_FONT_BASE, baf_FONT_SAFE_TEXT)
    baf_gFontNum = PickFont(ws, baf_FONT_NUM, baf_FONT_SAFE_NUM, baf_FONT_SAFE_TEXT)
    ws.Cells.Clear
End Sub

' Resolves the fonts once per session (on a scratch workbook)
Public Sub baf_EnsureFonts()
    Dim wbTmp As Workbook
    If baf_gFontBase <> "" And baf_gFontName <> "" And baf_gFontNum <> "" Then Exit Sub
    Set wbTmp = Workbooks.Add(xlWBATWorksheet)
    baf_ResolveFonts wbTmp.Worksheets(1)
    wbTmp.Close SaveChanges:=False
End Sub

' Name of the output folder (next to the tool)
Public Function baf_OutputFolderName() As String
    baf_OutputFolderName = baf_ArT("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A")
End Function

Private Function PickFont(ByVal ws As Worksheet, ParamArray names() As Variant) As String
    Dim i As Long, first As String, chosen As String
    first = CStr(names(LBound(names)))
    If Not baf_CHECK_FONTS Then
        PickFont = first
        Exit Function
    End If
    chosen = CStr(names(UBound(names)))
    For i = LBound(names) To UBound(names) - 1
        If FontAvailable(ws, CStr(names(i))) Then
            chosen = CStr(names(i))
            Exit For
        End If
    Next i
    PickFont = chosen
    If StrComp(chosen, first, vbTextCompare) <> 0 Then
        If InStr(baf_gFontNotes, first & " ->") = 0 Then baf_gFontNotes = baf_gFontNotes & vbLf & first & " -> " & chosen
    End If
End Function

Private Function FontAvailable(ByVal ws As Worksheet, ByVal fontName As String) As Boolean
    Dim samples(1 To 2) As String, i As Long, wBase As Double, wTest As Double
    samples(1) = "MWmwIl0123456789 The quick brown fox " & baf_ArT("627 644 639 631 628 64A 629 20 645 62F 631 633 629")
    samples(2) = "00.00 18.75 05.50 Notef11 MoyenneMe"
    On Error GoTo Unknown
    For i = 1 To 2
        wBase = MeasureWidth(ws, "Zz No Such Font Zz", samples(i))
        wTest = MeasureWidth(ws, fontName, samples(i))
        If Abs(wTest - wBase) > 0.3 Then
            FontAvailable = True
            Exit Function
        End If
    Next i
    Exit Function
Unknown:
    FontAvailable = True
End Function

Private Function MeasureWidth(ByVal ws As Worksheet, ByVal fontName As String, ByVal sample As String) As Double
    With ws.Range("A1")
        .Value2 = sample
        .Font.Name = fontName
        .Font.Size = 14
        .WrapText = False
    End With
    ws.Columns(1).AutoFit
    MeasureWidth = ws.Columns(1).ColumnWidth
End Function

'==============================================================================
' CELL STYLES (same look as modExamenAnalysis02)
'==============================================================================
Public Sub baf_StyleBanner(ByVal rng As Range)
    With rng
        .Font.Name = baf_gFontBase: .Font.Size = 16: .Font.Bold = True
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
End Sub

Public Sub baf_StyleTableTitle(ByVal rng As Range)
    With rng
        .Font.Name = baf_gFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub baf_StyleGrid(ByVal rng As Range)
    With rng
        .Font.Name = baf_gFontNum: .Font.Size = 15: .Font.Bold = False
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        .Rows.RowHeight = baf_DATA_ROW_HEIGHT
    End With
End Sub

Public Sub baf_StyleHead(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = baf_gFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub baf_StyleLabel(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = baf_gFontName: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub baf_MergeText(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal txt As String)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        If c2 > c1 Then .Merge
        .Value2 = txt
    End With
End Sub

Public Sub baf_PutLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    baf_MergeText ws, r, 1, 3, txt
    baf_StyleLabel ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
End Sub

'==============================================================================
' DISCOVER the Data sheet (subjects, term averages, sections)
'==============================================================================
Private Function HasMarks(ByVal lo As ListObject, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim n As Double
    On Error Resume Next
    n = Application.WorksheetFunction.CountIf(lo.ListColumns(c1).DataBodyRange, ">0") + _
        Application.WorksheetFunction.CountIf(lo.ListColumns(c2).DataBodyRange, ">0")
    HasMarks = (n > 0)
End Function

' Fills the g* variables. Subjects = header pairs "<name> <F> 1" / "<name> <F> 2" that
' contain at least one mark > 0 (a zero means "not taught / exempt"). Returns False + errText on trouble.
Public Function baf_DiscoverData(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    Dim c As Long, k As Long, c2 As Long, mx As Long
    Dim h As String, b As String, sfx1 As String, sfx2 As String, mk As String
    Dim rngSec As Range

    baf_DiscoverData = False: errText = ""
    On Error GoTo EH

    Set ws = wb.Worksheets(baf_DATA_SHEET)
    Set lo = ws.ListObjects(baf_DATA_TABLE)
    baf_gNCols = lo.ListColumns.Count
    baf_gRows = lo.ListRows.Count
    If baf_gRows < 1 Then errText = "Data table is empty": Exit Function

    ReDim baf_gHdr(1 To baf_gNCols)
    baf_gSecCol = 0
    For c = 1 To baf_gNCols
        baf_gHdr(c) = CStr(lo.ListColumns(c).Name)
        If baf_gSecCol = 0 Then
            If StrComp(baf_NormKey(baf_gHdr(c)), baf_ArT("627 644 642 633 645"), vbTextCompare) = 0 Then baf_gSecCol = c
        End If
    Next c
    If baf_gSecCol < 7 Then errText = "Section column not found in Data": Exit Function
    If baf_gSecCol + 4 > baf_gNCols Then errText = "Data table is missing its extra columns": Exit Function

    sfx1 = " " & ChrW(&H641) & " 1"
    sfx2 = " " & ChrW(&H641) & " 2"
    mk = baf_ArT("645 639 62F 644")

    baf_gSubjN = 0: baf_gAvgC1 = 0: baf_gAvgC2 = 0: baf_gPEIdx = 0
    ReDim baf_gSubjName(1 To baf_gSecCol)
    ReDim baf_gSubjC1(1 To baf_gSecCol)
    ReDim baf_gSubjC2(1 To baf_gSecCol)

    For c = 6 To baf_gSecCol - 1
        h = baf_NormKey(baf_gHdr(c))
        If Len(h) > 4 Then
            If Right$(h, 4) = sfx1 Then
                b = Left$(h, Len(h) - 4)
                c2 = 0
                For k = c + 1 To baf_gSecCol - 1
                    If StrComp(baf_NormKey(baf_gHdr(k)), b & sfx2, vbTextCompare) = 0 Then
                        c2 = k
                        Exit For
                    End If
                Next k
                If c2 > 0 Then
                    If HasMarks(lo, c, c2) Then
                        baf_gSubjN = baf_gSubjN + 1
                        baf_gSubjName(baf_gSubjN) = b
                        baf_gSubjC1(baf_gSubjN) = c
                        baf_gSubjC2(baf_gSubjN) = c2
                        If InStr(baf_NormAr(b), baf_ArT("627 644 628 62F 646 64A 629")) > 0 Then baf_gPEIdx = baf_gSubjN
                    End If
                End If
            ElseIf Left$(h, 4) = mk Then
                If Right$(h, 1) = "1" Then baf_gAvgC1 = c
                If Right$(h, 1) = "2" Then baf_gAvgC2 = c
            End If
        End If
    Next c

    If baf_gSubjN < 1 Then errText = "No subject with marks was found in Data": Exit Function
    If baf_gAvgC1 = 0 Or baf_gAvgC2 = 0 Then errText = "Term average columns were not found in Data": Exit Function

    ' sections that really have students
    Set rngSec = lo.ListColumns(baf_gSecCol).DataBodyRange
    mx = CLng(Application.WorksheetFunction.Max(rngSec))
    If mx < 1 Then errText = "Section numbers are missing in Data": Exit Function
    If mx > 60 Then errText = "Section number out of range": Exit Function
    baf_gSecN = 0
    ReDim baf_gSecList(1 To mx)
    For k = 1 To mx
        If Application.WorksheetFunction.CountIf(rngSec, k) > 0 Then
            baf_gSecN = baf_gSecN + 1
            baf_gSecList(baf_gSecN) = k
        End If
    Next k

    baf_gClassTitle = CStr(ws.Range("A5").Value)
    baf_DiscoverData = True
    Exit Function
EH:
    errText = "baf_DiscoverData: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function

'==============================================================================
' SHARED SHEET HELPERS (Analysis and Compare sheets)
'==============================================================================
' Cell-value rules: text / empty cells are never coloured (text counts as greater than any number)
Public Sub baf_ApplyPassColors(ByVal rng As Range, ByVal limitF As String, ByVal topF As String)
    Dim fc As Object
    On Error Resume Next
    rng.FormatConditions.Delete
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=limitF)
    fc.Interior.Color = RGB(192, 0, 0)
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:=limitF, Formula2:=topF)
    fc.Interior.Color = RGB(0, 128, 64)
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True
    Err.Clear
End Sub

Public Sub baf_PutStrip(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, _
                     ByVal cv As Long, ByVal lbl As String, ByVal fml As String)
    baf_MergeText ws, r, c1, c2, lbl
    baf_StyleHead ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    With ws.Cells(r, cv)
        .Formula = fml
        .NumberFormat = baf_FMT_COUNT
        .Font.Name = baf_gFontNum: .Font.Size = 15: .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    ws.Rows(r).RowHeight = baf_DATA_ROW_HEIGHT
End Sub

' header row of an analysis table: merged label in A:C, subject names rotated
Public Sub baf_PutHeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, _
                         ByVal lbl As String, ByRef hdrText() As String)
    Dim c As Long
    ws.Range(ws.Cells(r, baf_FIRST_NUM_COL), ws.Cells(r, lastCol)).NumberFormat = "@"
    For c = baf_FIRST_NUM_COL To lastCol
        ws.Cells(r, c).Value2 = hdrText(c)
    Next c
    baf_StyleHead ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
    ws.Range(ws.Cells(r, baf_FIRST_NUM_COL), ws.Cells(r, lastCol)).Orientation = 90
    baf_MergeText ws, r, 1, 3, lbl
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Size = 16
    ws.Rows(r).RowHeight = baf_HDR_ROW_HEIGHT
End Sub

Public Sub baf_SetColumnWidths(ByVal ws As Worksheet, ByVal lastCol As Long)
    ws.Columns(1).ColumnWidth = baf_ID_COL_WIDTH
    ws.Columns(2).ColumnWidth = 22
    ws.Columns(3).ColumnWidth = baf_DATE_COL_MIN_WIDTH
    ws.Range(ws.Columns(baf_FIRST_NUM_COL), ws.Columns(lastCol)).ColumnWidth = baf_NUM_COL_WIDTH
End Sub

Public Sub baf_SetupPrint(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True
    Err.Clear
End Sub

' After the full calculation: columns D -> last data column = AutoFit (Analysis01/02 and Compare)
Public Sub baf_AutoFitOutput(ByVal wb As Workbook)
    Dim ws As Worksheet, nm As Variant, lastCol As Long
    lastCol = baf_FIRST_NUM_COL + baf_gSubjN
    On Error Resume Next
    For Each nm In Array(baf_ANALYSIS_PREFIX & "01", baf_ANALYSIS_PREFIX & "02", baf_COMPARE_SHEET)
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(nm))
        If Not ws Is Nothing Then ws.Range(ws.Columns(baf_FIRST_NUM_COL), ws.Columns(lastCol)).EntireColumn.AutoFit
    Next nm
    Err.Clear
End Sub

' Protects Analysis01/02, Compare and Cart_Tec (no password). Returns "" when all are protected.
Public Function baf_ProtectOutput(ByVal wb As Workbook) As String
    Dim ws As Worksheet, nm As Variant
    For Each nm In Array(baf_ANALYSIS_PREFIX & "01", baf_ANALYSIS_PREFIX & "02", baf_COMPARE_SHEET, baf_CART_SHEET)
        On Error Resume Next
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(nm))
        On Error GoTo 0
        If ws Is Nothing Then
            baf_ProtectOutput = baf_ProtectOutput & "sheet not found: " & CStr(nm) & "; "
        Else
            On Error Resume Next
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
            If Err.Number <> 0 Then baf_ProtectOutput = baf_ProtectOutput & CStr(nm) & ": " & Err.Description & "; ": Err.Clear
            On Error GoTo 0
        End If
    Next nm
End Function
