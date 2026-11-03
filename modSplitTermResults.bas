Attribute VB_Name = "modSplitTermResults"
Option Explicit

' Split Term Results : ONE action that creates BOTH files from the raw \Source02\ files.
'   FirstTermOnly.xlsx  = first column of every pair (term 1), second column omitted
'   SecondTermOnly.xlsx = second column of every pair (term 2), first column omitted
' Source  : \Source02\ is collected ONCE, in memory, by the embedded engine (only read).
' Output  : folder next to the tool workbook; a file that already exists is never replaced
'           (<name>.xlsx, <name>_01.xlsx, ...). Every sheet is protected (no password;
'           filtering stays allowed). Ministry Term 02 exports hold two columns per subject.
Private Const GROUP_SIZE As Long = 2
Private Const FIRST_SUBJECT_COL As Long = 6

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

Public Sub SplitTermResults()
    Dim wbSource As Workbook
    Dim srcFolder As String, outDir As String
    Dim srcCopied As Long, srcSkipped As Long, srcNotes As String
    Dim report As String, fatalMsg As String, msg As String
    Dim term As Long, okCount As Long
    Dim prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean
    Dim prevSec As MsoAutomationSecurity
    Dim settingsCaptured As Boolean

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save this workbook before running the macro.", vbExclamation, "Split Term Results": Exit Sub
    End If
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Use a local or locally synced folder.", vbExclamation, "Split Term Results": Exit Sub
    End If
    srcFolder = ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME & "\"
    outDir = ThisWorkbook.Path & "\" & OutputFolderName() & "\"

    On Error GoTo FatalError
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevSec = Application.AutomationSecurity
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    ' Raw source files are collected ONCE in memory (the source files are only read).
    If Not SrcBuildWorkbook(srcFolder, wbSource, srcCopied, srcSkipped, srcNotes) Then
        fatalMsg = "No usable source file was found in " & srcFolder
        GoTo CleanExit
    End If
    If Not EnsureFolder(outDir) Then
        fatalMsg = "The output folder could not be created: " & outDir
        GoTo CleanExit
    End If

    For term = 1 To 2
        If ExportTerm(wbSource, term, outDir, report) Then okCount = okCount + 1
    Next term
    GoTo CleanExit

FatalError:
    fatalMsg = CStr(Err.Number) & " : " & Err.Description
    On Error GoTo 0
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If settingsCaptured Then
        Application.AutomationSecurity = prevSec
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
    End If
    On Error GoTo 0

    msg = "Created: " & okCount & " / 2" & vbCrLf & "Source files read: " & srcCopied & _
          "   skipped: " & srcSkipped & vbCrLf & vbCrLf & report
    If Len(srcNotes) > 0 Then msg = msg & vbCrLf & "Source notes:" & vbCrLf & srcNotes
    If Len(fatalMsg) > 0 Then msg = msg & vbCrLf & "ERROR: " & fatalMsg
    MsgBox msg, IIf(okCount = 2 And Len(fatalMsg) = 0, vbInformation, vbExclamation), "Split Term Results"
End Sub

' Builds, protects and saves ONE term file. Returns True when the file was saved.
' Nothing is saved when a sheet failed; a file that already exists is never replaced.
Private Function ExportTerm(ByVal wbSource As Workbook, ByVal term As Long, ByVal outDir As String, _
                            ByRef report As String) As Boolean
    Dim wbResult As Workbook, wsSource As Worksheet, placeholder As Worksheet, ws As Worksheet
    Dim tableIndex As Long, status As Long
    Dim total As Long, okSheets As Long, skippedSheets As Long, failedSheets As Long
    Dim reason As String, details As String, resultFilePath As String, label As String

    label = OutBaseName(term)
    On Error GoTo EH
    Set wbResult = Workbooks.Add(xlWBATWorksheet)
    Set placeholder = wbResult.Worksheets(1)
    For Each wsSource In wbSource.Worksheets
        total = total + 1
        status = ConvertSheet(wsSource, wbResult, tableIndex, reason, term)
        Select Case status
            Case 1
                okSheets = okSheets + 1
            Case 0
                skippedSheets = skippedSheets + 1
                details = details & "  - " & wsSource.Name & " : no student data / no subject columns" & vbCrLf
            Case Else
                failedSheets = failedSheets + 1
                details = details & "  - " & wsSource.Name & " : " & reason & vbCrLf
        End Select
    Next wsSource

    If okSheets > 0 And failedSheets = 0 Then
        placeholder.Delete
        Set placeholder = Nothing
        For Each ws In wbResult.Worksheets
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
        Next ws
        resultFilePath = NextFreePath(outDir, label, ".xlsx")
        wbResult.SaveAs Filename:=resultFilePath, FileFormat:=xlOpenXMLWorkbook
        ExportTerm = True
        report = report & label & " : saved  " & resultFilePath & vbCrLf
    Else
        report = report & label & " : NOT created" & vbCrLf
    End If
    report = report & "    sheets " & total & ", converted " & okSheets & ", skipped " & skippedSheets & _
             ", failed " & failedSheets & vbCrLf & details
    GoTo Done

EH:
    ExportTerm = False
    report = report & label & " : ERROR " & Err.Number & " : " & Err.Description & vbCrLf
    Err.Clear

Done:
    On Error Resume Next
    If Not wbResult Is Nothing Then wbResult.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function OutBaseName(ByVal term As Long) As String
    If term = 1 Then OutBaseName = "FirstTermOnly" Else OutBaseName = "SecondTermOnly"
End Function

Private Sub WriteTermTitle(ByVal target As Range, ByVal title As String, ByVal source As Range)
    With target
        .UnMerge
        .ClearContents
        .Merge
        .Value = title
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = True
        .Font.Name = source.Font.Name
        .Font.Size = 11
        .Font.Bold = True
    End With
End Sub

' Returns 1 = converted, 0 = skipped, -1 = failed
Private Function ConvertSheet(ByVal wsSource As Worksheet, ByVal wbResult As Workbook, _
                              ByRef tableIndex As Long, _
                              ByRef reasonOut As String, _
                              ByVal term As Long) As Long

    Dim wsResult As Worksheet, foundRange As Range
    Dim LastRow As Long, LastCol As Long
    Dim srcData As Variant, resData As Variant
    Dim SourceCol As Long, ResultCol As Long, RowNum As Long, c As Long
    Dim resRows As Long, resCols As Long, newCols As Long, off As Long
    Dim sheetName As String
    Dim tbl As ListObject, tblRange As Range
    Dim seenHeaders As Object, h As Long, headerText As String, genCounter As Long

    reasonOut = ""
    ConvertSheet = 0
    off = term - 1            ' 0 = first column of each pair, 1 = second column
    sheetName = wsSource.Name

    On Error GoTo SheetErr

    LastCol = wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column
    Set foundRange = wsSource.Range(wsSource.Cells(3, 1), wsSource.Cells(wsSource.Rows.Count, LastCol)).Find( _
        What:="*", After:=wsSource.Cells(3, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not foundRange Is Nothing Then LastRow = foundRange.Row

    If LastRow < 3 Or LastCol < FIRST_SUBJECT_COL Then
        Exit Function
    End If

    Set wsResult = wbResult.Worksheets.Add(After:=wbResult.Worksheets(wbResult.Worksheets.Count))

    On Error Resume Next
    wsResult.Name = Left(SanitizeSheetName(sheetName), 31)
    If Err.Number <> 0 Then wsResult.Name = "Sheet" & wbResult.Worksheets.Count: Err.Clear
    On Error GoTo SheetErr

    wsResult.DisplayRightToLeft = True
    srcData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(LastRow, LastCol)).Value

    resRows = LastRow - 1
    If term = 1 Then
        newCols = (LastCol - (FIRST_SUBJECT_COL - 1) + GROUP_SIZE - 1) \ GROUP_SIZE
    Else
        newCols = (LastCol - (FIRST_SUBJECT_COL - 1)) \ GROUP_SIZE
    End If
    resCols = (FIRST_SUBJECT_COL - 1) + newCols
    ReDim resData(1 To resRows, 1 To resCols)

    For RowNum = 1 To resRows
        For ResultCol = 1 To FIRST_SUBJECT_COL - 1
            resData(RowNum, ResultCol) = srcData(RowNum, ResultCol)
        Next ResultCol
    Next RowNum

    ' Header Processing: Strip the term suffix (e.g., "\u0641 1")
    ResultCol = FIRST_SUBJECT_COL
    For SourceCol = FIRST_SUBJECT_COL + off To LastCol Step GROUP_SIZE
        resData(1, ResultCol) = StripTermSuffix(srcData(1, SourceCol), term)
        ResultCol = ResultCol + 1
    Next SourceCol

    ' Copy the term values unchanged, including blanks, text and Excel errors.
    For RowNum = 2 To resRows
        ResultCol = FIRST_SUBJECT_COL
        For SourceCol = FIRST_SUBJECT_COL + off To LastCol Step GROUP_SIZE
            resData(RowNum, ResultCol) = srcData(RowNum, SourceCol)
            ResultCol = ResultCol + 1
        Next SourceCol
    Next RowNum

    ' Set text/date/score formats before values are written.
    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        .Columns("D:E").NumberFormat = "@"
        .Range(.Columns(6), .Columns(resCols)).NumberFormat = "0.00"
    End With

    wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(LastRow, resCols)).Value = resData

    ' Unique, non-empty headers (case-insensitive)
    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    genCounter = 0
    For h = 1 To resCols
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

    Set tblRange = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(LastRow, resCols))
    Set tbl = wsResult.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "Data" & Format(tableIndex + 1, "00")
    tbl.TableStyle = ""
    tbl.ShowTableStyleColumnStripes = False
    tbl.ShowTableStyleRowStripes = False
    tbl.ShowTableStyleFirstColumn = False
    tbl.ShowTableStyleLastColumn = False
    tbl.ShowAutoFilter = True

    ' Map formatting to the reduced output columns without copying whole rows.
    For c = 1 To 5
        With wsResult.Range(wsResult.Cells(2, c), wsResult.Cells(LastRow, c))
            .Font.Name = wsSource.Cells(3, c).Font.Name
            .Font.Size = wsSource.Cells(3, c).Font.Size
        End With
    Next c
    ResultCol = FIRST_SUBJECT_COL
    For SourceCol = FIRST_SUBJECT_COL + off To LastCol Step GROUP_SIZE
        With wsResult.Range(wsResult.Cells(2, ResultCol), wsResult.Cells(LastRow, ResultCol))
            .Font.Name = wsSource.Cells(3, SourceCol).Font.Name
            .Font.Size = wsSource.Cells(3, SourceCol).Font.Size
        End With
        ResultCol = ResultCol + 1
    Next SourceCol
    With wsResult
        With .Range(.Cells(2, 1), .Cells(LastRow, resCols))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With
        .Columns("A:E").AutoFit
        .Range(.Columns(FIRST_SUBJECT_COL), .Columns(resCols)).ColumnWidth = 5.7
        With .Range(.Cells(2, 4), .Cells(2, resCols))
            .Orientation = xlUpward
            .Font.Bold = True
        End With
        .Rows(2).RowHeight = wsSource.Rows(2).RowHeight
        .Rows(1).RowHeight = 30
    End With
    WriteTermTitle wsResult.Range("B1:F1"), TermTitle(CStr(wsSource.Range("B1").Value), term), wsSource.Range("B1")
    WriteTermTitle wsResult.Range("H1:Q1"), CStr(wsSource.Range("H1").Value), wsSource.Range("H1")

    tableIndex = tableIndex + 1
    ConvertSheet = 1
    Exit Function

SheetErr:
    reasonOut = CStr(Err.Number) & " : " & Err.Description
    ConvertSheet = -1
    On Error Resume Next
    If Not wsResult Is Nothing Then wsResult.Delete
    Err.Clear
    On Error GoTo 0
End Function

' ==============================================================================
' Removes the term suffix (e.g., "\u0641 1", "\u0641 2") from headers using RegExp.
' Handles Arabic "\u0641", Western digits, Arabic-Indic, and Persian digits.
' ==============================================================================
Private Function StripTermSuffix(ByVal v As Variant, ByVal term As Long) As Variant
    Dim re As Object, t As String, r As String

    If IsError(v) Then StripTermSuffix = v: Exit Function
    If IsEmpty(v) Then StripTermSuffix = v: Exit Function

    t = CStr(v)
    On Error GoTo Fail
    Set re = CreateObject("VBScript.RegExp")
    
    ' Matches: optional spaces + Arabic "\u0641" (ChrW 1601) + optional spaces + digits at the end
    re.Pattern = "[\s\u00A0]*" & ChrW(1601) & "[\s\u00A0]*[" & CStr(term) & ChrW(&H660 + term) & ChrW(&H6F0 + term) & "][\s\u00A0]*$"
    re.Global = False
    
    r = Trim(re.Replace(t, ""))
    If Len(r) = 0 Then r = Trim(t)      ' Fallback to original if it becomes empty
    StripTermSuffix = r
    Exit Function
Fail:
    Err.Clear
    StripTermSuffix = v
End Function

Private Function SanitizeSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad): s = Replace(s, bad(i), "_"): Next i
    SanitizeSheetName = s
End Function

' Replace the fixed title; preserve the matched school year verbatim.
Private Function TermTitle(ByVal sourceTitle As String, ByVal term As Long) As String
    Dim rx As Object, matches As Object, digits As String, fixedTitle As String
    Set rx = CreateObject("VBScript.RegExp")
    digits = "[0-9" & ChrW(&H660) & "-" & ChrW(&H669) & ChrW(&H6F0) & "-" & ChrW(&H6F9) & "]"
    rx.Pattern = digits & "{4}\s*[-/" & ChrW(&H2013) & ChrW(&H2014) & "]\s*" & digits & "{4}"
    rx.Global = False
    Set matches = rx.Execute(sourceTitle)
    If matches.Count = 0 Then
        Err.Raise vbObjectError + 720, "TermTitle", "School year not found in source B1."
    End If
    fixedTitle = ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1574) & ChrW(1580) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1601) & ChrW(1589) & ChrW(1604) & ChrW(32) & ChrW(1575) & ChrW(1604)
    If term = 1 Then
        fixedTitle = fixedTitle & ChrW(1571) & ChrW(1608) & ChrW(1604)
    Else
        fixedTitle = fixedTitle & ChrW(1579) & ChrW(1575) & ChrW(1606) & ChrW(1610)
    End If
    TermTitle = fixedTitle & " " & matches(0).Value
End Function


' ==============================================================================
' OUTPUT FOLDER + FREE FILE NAME (nothing that already exists is replaced)
' ==============================================================================
' Name of the output folder created next to the tool workbook (hex code points)
Private Function OutputFolderName() As String
    OutputFolderName = AR("062A 062D 0644 064A 0644 0020 0646 062A 0627 0626 062C 0020 0627 0644 0641 0635 0644 0020 0627 0644 062B 0627 0646 064A")
End Function

' stem.xlsx, then stem_01.xlsx, stem_02.xlsx ... (folder must end with a backslash)
Private Function NextFreePath(ByVal folder As String, ByVal stem As String, ByVal ext As String) As String
    Dim p As String, n As Long
    p = folder & stem & ext
    Do While FileExistsPath(p) Or IsOpenInExcel(p)
        n = n + 1
        p = folder & stem & "_" & Format$(n, "00") & ext
    Loop
    NextFreePath = p
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
