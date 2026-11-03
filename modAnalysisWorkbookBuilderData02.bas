Attribute VB_Name = "modBiuldAnalysisfilesData02"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesData02   (replaces the "ResultFlow02 -> Data" step)
' Purpose : Reads \Source02\ with the SAME embedded collection engine as
'           ResultFlow02, groups the sections exactly as the Map of ResultFlow02
'           (12 files from 17 source sheets) and writes the Data sheet of each
'           NEW workbook. Nothing is distributed into existing files any more.
' Notes   : - Pure ASCII: Arabic text = hex code points (see AR).
'           - The source files are only read: nothing is converted or moved.
'           - Engine = lines of ResultFlow02 (SortedNames ... end), unchanged
'             except one block in SrcProcessFile that keeps rows A1:A4 of the
'             source (identity lines) in cell A1 of each collected sheet.
'==============================================================================

Private Const SOURCE_FOLDER_NAME As String = "Source02"

' ---- Collection engine constants (embedded copy of modCollector - keep identical) ----
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

    ' Keep the identity lines (A1:A4 of the source) for the Data sheet of the new files
    On Error Resume Next
    wsResult.Range("A1").Value = CleanText(CStr(wsSource.Range("A1").Value)) & "|" & _
                                 CleanText(CStr(wsSource.Range("A2").Value)) & "|" & _
                                 CleanText(CStr(wsSource.Range("A3").Value)) & "|" & _
                                 CleanText(CStr(wsSource.Range("A4").Value))
    Err.Clear
    On Error GoTo EH

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


' ==============================================================================
' PUBLIC API (new in modBiuldAnalysisfilesData02)
' ==============================================================================

' Builds the collected workbook in memory from \Source02\ and indexes its sheets by H1.
' wbGlobal stays OPEN and unsaved: the caller closes it with SaveChanges:=False.
Public Function baf_LoadSource(ByRef wbGlobal As Workbook, ByRef sheetIndex As Object, _
                           ByRef dupKeys As Object, ByRef report As String) As Boolean
    Dim srcFolder As String, copied As Long, skipped As Long, lst As String
    Dim ws As Worksheet, k As String

    baf_LoadSource = False
    report = ""

    If Len(ThisWorkbook.Path) = 0 Then
        report = "  - Save the tool workbook first (its folder is needed to find " & SOURCE_FOLDER_NAME & ")." & vbCrLf
        Exit Function
    End If
    If Not FolderExistsPath(ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME) Then
        report = "  - " & SOURCE_FOLDER_NAME & " folder not found next to the tool." & vbCrLf
        Exit Function
    End If

    srcFolder = ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME & "\"
    Set wbGlobal = Nothing
    If Not SrcBuildWorkbook(srcFolder, wbGlobal, copied, skipped, lst) Then
        report = lst
        Exit Function
    End If
    report = lst

    Set sheetIndex = CreateObject("Scripting.Dictionary")
    sheetIndex.CompareMode = vbTextCompare
    Set dupKeys = CreateObject("Scripting.Dictionary")
    dupKeys.CompareMode = vbTextCompare

    For Each ws In wbGlobal.Worksheets
        k = NormalizeKey(CellText(ws.Range("H1")))
        If k <> "" Then
            If sheetIndex.Exists(k) Then
                If Not dupKeys.Exists(k) Then dupKeys.Add k, True
            Else
                sheetIndex.Add k, ws
            End If
        End If
    Next ws

    baf_LoadSource = True
End Function

' stem -> array of section titles (the H1 text of each source sheet that feeds the file).
' Same grouping as the Map of ResultFlow02: 12 files from 17 source sheets.
Public Function baf_GetFileMap() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    AddGroup m, "1LitAnalysis02", AR("623 648 644 649 20 62B 627 646 648 64A 20 62C 62F 639 20 645 634 62A 631 643 20 622 62F 627 628"), 1
    AddGroup m, "1TecnoGlobalAnalysis02", AR("623 648 644 649 20 62B 627 646 648 64A 20 62C 62F 639 20 645 634 62A 631 643 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627"), 4
    AddGroup m, "2GEAnalysis02", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F"), 1
    AddGroup m, "2LEAnalysis02", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 644 63A 627 62A 20 623 62C 646 628 64A 629 20 623 644 645 627 646 64A 629"), 1
    AddGroup m, "2MathAnalysis02", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 631 64A 627 636 64A 627 62A"), 1
    AddGroup m, "2PhiloAnalysis02", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 622 62F 627 628 20 648 641 644 633 641 629"), 1
    AddGroup m, "2SEGlobalAnalysis02", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 639 644 648 645 20 62A 62C 631 64A 628 64A 629"), 2
    AddGroup m, "3GEAnalysis02", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F"), 1
    AddGroup m, "3LEAnalysis02", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 644 63A 627 62A 20 623 62C 646 628 64A 629 20 623 644 645 627 646 64A 629"), 1
    AddGroup m, "3MathAnalysis02", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 631 64A 627 636 64A 627 62A"), 1
    AddGroup m, "3PhiloAnalysis02", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 622 62F 627 628 20 648 641 644 633 641 629"), 1
    AddGroup m, "3SEGlobalAnalysis02", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 639 644 648 645 20 62A 62C 631 64A 628 64A 629"), 2
    Set baf_GetFileMap = m
End Function

Private Sub AddGroup(ByVal m As Object, ByVal stem As String, ByVal prefix As String, ByVal cnt As Long)
    Dim a() As String, i As Long
    ReDim a(0 To cnt - 1)
    For i = 1 To cnt
        a(i - 1) = prefix & " " & Format$(i, "00")
    Next i
    m.Add stem, a
End Sub

' True when every section of the group exists once and has data
Public Function baf_GroupReady(ByVal sheetIndex As Object, ByVal dupKeys As Object, _
                           ByVal titles As Variant, ByRef why As String) As Boolean
    Dim i As Long, k As String, ok As Boolean, wsT As Worksheet
    ok = True
    why = ""
    For i = LBound(titles) To UBound(titles)
        k = NormalizeKey(CStr(titles(i)))
        If Not sheetIndex.Exists(k) Then
            ok = False: why = why & "    missing section: " & k & vbCrLf
        ElseIf dupKeys.Exists(k) Then
            ok = False: why = why & "    duplicate section: " & k & vbCrLf
        Else
            Set wsT = sheetIndex(k)
            If LastUsedRow(wsT) < 3 Then
                ok = False: why = why & "    empty section: " & k & vbCrLf
            End If
        End If
    Next i
    baf_GroupReady = ok
End Function

' "<class title> 01"  ->  "<class title>"
Private Function ClassTitleOf(ByVal key As String) As String
    key = NormalizeKey(key)
    If Len(key) > 3 Then
        If Mid$(key, Len(key) - 2, 1) = " " And IsNumeric(Right$(key, 2)) Then key = Left$(key, Len(key) - 3)
    End If
    ClassTitleOf = key
End Function

' ==============================================================================
' DATA SHEET BUILDER
'   Row 1-4 : identity lines taken from the source file (A1:A4)
'   Row 5   : class title
'   Row 6   : headers (source headers unchanged) + extra columns
'   Row 7.. : all sections one after the other (the section number goes in the section column)
'   Extra columns : section | average of the two terms (formula) | transition/acceptance
'                   average (manual) | picture (manual) | notes (manual)
' ==============================================================================
Public Function baf_BuildDataSheet(ByVal wbNew As Workbook, ByVal sheetIndex As Object, _
                               ByVal titles As Variant, ByVal acceptHdr As String, _
                               ByRef errText As String) As Boolean
    Dim ws As Worksheet, wsSrc As Worksheet, wsFirst As Worksheet
    Dim nSec As Long, i As Long, c As Long, r As Long, rr As Long
    Dim W As Long, total As Long, lastRow As Long, lastCol As Long
    Dim rowsOf() As Long, secVal() As Long, hdrBase() As String
    Dim key As String, classTitle As String, txt As String, h As String, nm As String
    Dim arr As Variant, outArr() As Variant, hdrRow() As Variant, extras As Variant
    Dim seen As Object, gen As Long
    Dim lo As ListObject, lines() As String
    Dim lastC As Long, a1 As Long, a2 As Long, fRow As Long, lRow As Long
    Dim mk As String, secStart As Long

    baf_BuildDataSheet = False
    errText = ""
    On Error GoTo EH

    nSec = UBound(titles) - LBound(titles) + 1
    ReDim rowsOf(1 To nSec)
    ReDim secVal(1 To nSec)

    ' ---- pass 1: validate sections, count rows
    For i = 1 To nSec
        key = NormalizeKey(CStr(titles(LBound(titles) + i - 1)))
        Set wsSrc = sheetIndex(key)
        lastCol = wsSrc.Cells(2, wsSrc.Columns.Count).End(xlToLeft).Column
        lastRow = LastUsedRow(wsSrc)
        rowsOf(i) = lastRow - 2
        If rowsOf(i) < 1 Then Err.Raise vbObjectError + 801, "baf_BuildDataSheet", "Empty section: " & key
        If i = 1 Then
            Set wsFirst = wsSrc
            W = lastCol
            If W < 7 Then Err.Raise vbObjectError + 804, "baf_BuildDataSheet", "Too few columns in: " & key
            ReDim hdrBase(1 To W)
            For c = 1 To W
                hdrBase(c) = NormalizeKey(CellText(wsSrc.Cells(2, c)))
            Next c
            classTitle = ClassTitleOf(key)
        Else
            If lastCol <> W Then Err.Raise vbObjectError + 802, "baf_BuildDataSheet", "Column count differs in: " & key
            For c = 1 To W
                If StrComp(NormalizeKey(CellText(wsSrc.Cells(2, c))), hdrBase(c), vbTextCompare) <> 0 Then
                    Err.Raise vbObjectError + 803, "baf_BuildDataSheet", "Header differs (column " & c & ") in: " & key
                End If
            Next c
        End If
        If IsNumeric(Right$(key, 2)) Then secVal(i) = CLng(Right$(key, 2)) Else secVal(i) = i
        total = total + rowsOf(i)
    Next i

    ' ---- headers: unique and non-empty
    lastC = W + 5
    ReDim hdrRow(1 To lastC)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For c = 1 To W
        h = hdrBase(c)
        If h = "" Or seen.Exists(h) Then
            Do
                gen = gen + 1
                h = "Column" & gen
            Loop While seen.Exists(h)
        End If
        seen.Add h, True
        hdrRow(c) = h
    Next c

    extras = Array(AR("627 644 642 633 645"), AR("645 639 62F 644 20 627 644 641 635 644 64A 646"), acceptHdr, AR("627 644 635 648 631 629"), AR("627 644 645 644 627 62D 638 627 62A"))
    For i = 0 To 4
        nm = CStr(extras(i))
        Do While seen.Exists(nm)
            nm = nm & " 2"
        Loop
        seen.Add nm, True
        hdrRow(W + 1 + i) = nm
    Next i

    ' ---- pass 2: gather the marks
    ReDim outArr(1 To total, 1 To W)
    rr = 0
    For i = 1 To nSec
        key = NormalizeKey(CStr(titles(LBound(titles) + i - 1)))
        Set wsSrc = sheetIndex(key)
        arr = wsSrc.Range(wsSrc.Cells(3, 1), wsSrc.Cells(2 + rowsOf(i), W)).Value
        For r = 1 To rowsOf(i)
            rr = rr + 1
            For c = 1 To W
                outArr(rr, c) = arr(r, c)
            Next c
        Next r
    Next i

    ' ---- write
    Set ws = wbNew.Worksheets(1)
    ws.Name = baf_DATA_SHEET
    ws.DisplayRightToLeft = True
    fRow = baf_DATA_FIRST_ROW
    lRow = baf_DATA_FIRST_ROW + total - 1

    With ws
        .Range(.Cells(baf_DATA_HDR_ROW, 1), .Cells(baf_DATA_HDR_ROW, lastC)).NumberFormat = "@"
        .Range(.Cells(baf_DATA_HDR_ROW, 1), .Cells(baf_DATA_HDR_ROW, lastC)).Value = hdrRow

        ' number formats BEFORE the data so that text columns stay text
        .Range(.Cells(fRow, 1), .Cells(lRow, 1)).NumberFormat = baf_FMT_COUNT
        .Range(.Cells(fRow, 2), .Cells(lRow, 2)).NumberFormat = "@"
        .Range(.Cells(fRow, 3), .Cells(lRow, 3)).NumberFormat = "yyyy-mm-dd"
        .Range(.Cells(fRow, 4), .Cells(lRow, 5)).NumberFormat = "@"
        .Range(.Cells(fRow, 6), .Cells(lRow, W)).NumberFormat = baf_FMT_DEC
        .Range(.Cells(fRow, W + 1), .Cells(lRow, W + 1)).NumberFormat = baf_FMT_COUNT
        .Range(.Cells(fRow, W + 2), .Cells(lRow, W + 3)).NumberFormat = baf_FMT_DEC

        .Cells(fRow, 1).Resize(total, W).Value = outArr

        secStart = fRow
        For i = 1 To nSec
            .Range(.Cells(secStart, W + 1), .Cells(secStart + rowsOf(i) - 1, W + 1)).Value = secVal(i)
            secStart = secStart + rowsOf(i)
        Next i
    End With

    ' ---- table named Data (no style: the look is applied by the style helpers)
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(baf_DATA_HDR_ROW, 1), ws.Cells(lRow, lastC)), , xlYes)
    lo.Name = baf_DATA_TABLE
    lo.TableStyle = ""
    lo.ShowTableStyleRowStripes = False
    lo.ShowTableStyleColumnStripes = False
    lo.ShowTableStyleFirstColumn = False
    lo.ShowTableStyleLastColumn = False
    lo.ShowAutoFilter = True

    ' ---- average of the two terms = formula over the last two "average" columns
    mk = AR("645 639 62F 644")
    a1 = 0: a2 = 0
    For c = W To 6 Step -1
        If Left$(CStr(hdrRow(c)), 4) = mk Then
            If a2 = 0 Then
                a2 = c
            ElseIf a1 = 0 Then
                a1 = c
                Exit For
            End If
        End If
    Next c
    If a1 = 0 Or a2 = 0 Then a1 = W - 1: a2 = W
    ws.Range(ws.Cells(fRow, W + 2), ws.Cells(lRow, W + 2)).Formula = _
        "=IFERROR(AVERAGE(" & baf_ColLetter(a1) & fRow & ":" & baf_ColLetter(a2) & fRow & ")," & baf_QT & baf_QT & ")"

    ' ---- identity lines (rows 1-5)
    lines = Split(CellText(wsFirst.Range("A1")), "|")
    For r = 1 To 4
        txt = ""
        If UBound(lines) >= r - 1 Then txt = Trim$(lines(r - 1))
        baf_MergeText ws, r, 1, W, txt
    Next r
    baf_MergeText ws, 5, 1, W, classTitle

    ' ---- look
    With ws
        baf_StyleGrid .Range(.Cells(baf_DATA_HDR_ROW, 1), .Cells(lRow, lastC))
        baf_StyleHead .Range(.Cells(baf_DATA_HDR_ROW, 1), .Cells(baf_DATA_HDR_ROW, lastC))
        .Range(.Cells(baf_DATA_HDR_ROW, 4), .Cells(baf_DATA_HDR_ROW, W + 3)).Orientation = 90
        .Rows(baf_DATA_HDR_ROW).RowHeight = baf_HDR_ROW_HEIGHT

        With .Range(.Cells(fRow, 2), .Cells(lRow, 2))
            .Font.Name = baf_gFontName: .Font.Size = 14
            .HorizontalAlignment = xlRight
            .ShrinkToFit = True
        End With
        With .Range(.Cells(fRow, 4), .Cells(lRow, 5))
            .Font.Name = baf_gFontName: .Font.Size = 13
            .ShrinkToFit = True
        End With
        With .Range(.Cells(fRow, 3), .Cells(lRow, 3))
            .Font.Size = 14
        End With
        With .Range(.Cells(fRow, W + 5), .Cells(lRow, W + 5))
            .Font.Name = baf_gFontName: .Font.Size = 13
            .HorizontalAlignment = xlRight
            .WrapText = True
        End With

        .Columns(1).ColumnWidth = baf_ID_COL_WIDTH
        .Columns(2).ColumnWidth = 30
        .Columns(3).ColumnWidth = baf_DATE_COL_MIN_WIDTH + 1
        .Range(.Columns(4), .Columns(W + 3)).ColumnWidth = baf_NUM_COL_WIDTH
        .Columns(W + 4).ColumnWidth = 14
        .Columns(W + 5).ColumnWidth = 30

        For r = 1 To 4
            .Range(.Cells(r, 1), .Cells(r, W)).Font.Name = baf_gFontName
            .Range(.Cells(r, 1), .Cells(r, W)).Font.Size = 14
            .Range(.Cells(r, 1), .Cells(r, W)).HorizontalAlignment = xlCenter
            .Range(.Cells(r, 1), .Cells(r, W)).ReadingOrder = xlRTL
            .Rows(r).RowHeight = 24
        Next r
        With .Range(.Cells(5, 1), .Cells(5, W))
            .Font.Name = baf_gFontBase: .Font.Size = 16: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .ReadingOrder = xlRTL
        End With
        .Rows(5).RowHeight = 32
    End With

    baf_BuildDataSheet = True
    Exit Function
EH:
    errText = "baf_BuildDataSheet: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function
