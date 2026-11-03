Attribute VB_Name = "modBiuldAnalysisfilesBuildAll02"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesBuildAll02
' Purpose : ONE macro (Build_All) that creates all the Term 02 analysis files.
'           It is the action of the button "Second Term Analysis" of ControlPanel.
'
'   Build_All
'    |- modBiuldAnalysisfilesData02      -> Data       (collects \Source02\ and groups the sections as in the old Map)
'    |- modBiuldAnalysisfilesAnalysis02  -> Analysis01 / Analysis02
'    |- modBiuldAnalysisfilesCompare02   -> Compare    (term 1 against term 2)
'    |- modBiuldAnalysisfilesCartTec02   -> Cart_Tec
'
' Needs   : only the raw source files in \Source02\ (no Term02Results file, no previous button).
' Output  : \<output folder>\<file>.xlsx next to the tool (names of the old Map).
'           Nothing that already exists is replaced: a new run writes <file>_01.xlsx,
'           <file>_02.xlsx ... so the typed notes / pictures of earlier files are safe.
' Source  : \Source02\ is only read (never converted, moved or changed).
' Reports : the status sheet "Term02Status" (created by BuildControlPanel through
'           Status_CreateSheet) shows source folder, output folder, last run, duration,
'           success and the time of every file; a full text log is written in the output folder.
' Speed   : automatic calculation is off while a workbook is built and switched back on
'           (one full calculation) just before it is saved: the saved files are identical.
'==============================================================================

Private Const STATUS_SHEET As String = "Term02Status"
Private Const ROW_FIRST As Long = 12

Public Sub Build_All()
    Dim wbGlobal As Workbook
    Dim sheetIndex As Object, dupKeys As Object, map As Object
    Dim key As Variant, stem As String, titles As Variant
    Dim srcReport As String, why As String, errMsg As String, savedPath As String
    Dim outDir As String, logText As String, failBrief As String, logPath As String, msg As String
    Dim okCount As Long, failCount As Long, total As Long, idx As Long
    Dim oldSU As Boolean, oldEV As Boolean, oldDA As Boolean, oldSB As Variant, oldCalc As Long
    Dim t0 As Double, t1 As Double, secs As Double, secsAll As Double
    Dim firstErr As String, retried As Boolean

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the tool workbook first: the Source02 folder and the output folder are located next to it.", _
               vbExclamation, "Second Term Analysis"
        Exit Sub
    End If
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "The tool is on OneDrive/SharePoint (URL path). Move it to a local folder first.", _
               vbExclamation, "Second Term Analysis"
        Exit Sub
    End If

    oldSU = Application.ScreenUpdating
    oldEV = Application.EnableEvents
    oldDA = Application.DisplayAlerts
    oldSB = Application.StatusBar
    oldCalc = Application.Calculation
    t0 = Timer

    On Error GoTo Fatal
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set map = baf_GetFileMap()
    total = map.Count
    Status_Reset map

    Application.StatusBar = "Second Term Analysis : checking fonts ..."
    baf_EnsureFonts

    ' ---- source (collected once, in memory)
    Application.StatusBar = "Second Term Analysis : reading Source02 ..."
    If Not baf_LoadSource(wbGlobal, sheetIndex, dupKeys, srcReport) Then
        Status_Finish 0, total, 0, "Source02: " & Replace(Trim$(srcReport), vbCrLf, " | ")
        MsgBox "No usable source file was found." & vbCrLf & vbCrLf & Left$(srcReport, 700), vbCritical, "Second Term Analysis"
        GoTo Done
    End If

    ' ---- output folder (kept between runs)
    outDir = ThisWorkbook.Path & "\" & baf_OutputFolderName() & "\"
    If Not baf_EnsureFolderP(outDir) Then
        MsgBox "The output folder could not be created:" & vbCrLf & outDir, vbCritical, "Second Term Analysis"
        GoTo Done
    End If

    logText = "Second Term Analysis - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
              "Tool   : " & ThisWorkbook.FullName & vbCrLf & _
              "Output : " & outDir & vbCrLf & String$(70, "-") & vbCrLf

    ' ---- one file per group of the Map
    For Each key In map.Keys
        idx = idx + 1
        stem = CStr(key)
        titles = map(key)
        Application.StatusBar = "Second Term Analysis : " & idx & " / " & total & "  " & stem
        Status_Row idx, 1, "", 0
        t1 = Timer

        savedPath = ""
        retried = False
        If baf_GroupReady(sheetIndex, dupKeys, titles, why) Then
            errMsg = BuildOneFile(stem, titles, sheetIndex, outDir, savedPath)
            If errMsg <> "" Then
                ' one automatic retry: a one-off failure (file locked by a sync / antivirus
                ' program, Excel busy) usually passes the second time; a real data problem fails again
                firstErr = errMsg
                DoEvents
                savedPath = ""
                errMsg = BuildOneFile(stem, titles, sheetIndex, outDir, savedPath)
                If errMsg = "" Then retried = True
            End If
        Else
            errMsg = "SKIPPED - the source sections are not complete:" & vbCrLf & why
        End If

        secs = Timer - t1
        If secs < 0 Then secs = secs + 86400

        If errMsg = "" Then
            okCount = okCount + 1
            If retried Then
                Status_Row idx, 2, Mid$(savedPath, InStrRev(savedPath, "\") + 1) & "   [retry]", secs
                logText = logText & "OK     " & stem & "  ->  " & Mid$(savedPath, InStrRev(savedPath, "\") + 1) & _
                          "  (" & Format$(secs, "0.0") & " s)  after ONE RETRY" & vbCrLf & "       first error: " & _
                          Replace(Trim$(firstErr), vbCrLf, " | ") & vbCrLf
            Else
                Status_Row idx, 2, Mid$(savedPath, InStrRev(savedPath, "\") + 1), secs
                logText = logText & "OK     " & stem & "  ->  " & Mid$(savedPath, InStrRev(savedPath, "\") + 1) & _
                          "  (" & Format$(secs, "0.0") & " s)" & vbCrLf
            End If
        Else
            failCount = failCount + 1
            Status_Row idx, 3, Replace(Trim$(errMsg), vbCrLf, " | "), secs
            logText = logText & "FAILED " & stem & "  (" & Format$(secs, "0.0") & " s)" & vbCrLf & "       " & _
                      Replace(Trim$(errMsg), vbCrLf, vbCrLf & "       ") & vbCrLf
            failBrief = failBrief & "  - " & stem & " : " & Left$(Replace(Trim$(errMsg), vbCrLf, " "), 90) & vbCrLf
        End If
    Next key

    secsAll = Timer - t0
    If secsAll < 0 Then secsAll = secsAll + 86400

    If Trim$(srcReport) <> "" Then logText = logText & String$(70, "-") & vbCrLf & "Source notes:" & vbCrLf & srcReport
    If baf_gFontNotes <> "" Then logText = logText & String$(70, "-") & vbCrLf & "Fonts replaced:" & baf_gFontNotes & vbCrLf
    logText = logText & String$(70, "-") & vbCrLf & "Created " & okCount & " / " & total & " in " & Format$(secsAll, "0") & " s" & vbCrLf

    logPath = baf_NextFreePath(outDir, "BuildLog_" & Format$(Now, "yyyymmdd_hhnnss"), ".txt")
    WriteLog logPath, logText

    Status_Finish okCount, total, secsAll, ""

    msg = "Created : " & okCount & " / " & total & "   (" & Format$(secsAll, "0") & " s)" & vbCrLf & "Folder  : " & outDir
    If failCount > 0 Then msg = msg & vbCrLf & vbCrLf & "NOT created (" & failCount & "):" & vbCrLf & failBrief
    msg = msg & vbCrLf & "Full details: sheet " & STATUS_SHEET & " + " & Mid$(logPath, InStrRev(logPath, "\") + 1)
    MsgBox Left$(msg, 1000), IIf(failCount = 0, vbInformation, vbExclamation), "Second Term Analysis"

Done:
    On Error Resume Next
    If Not wbGlobal Is Nothing Then wbGlobal.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.StatusBar = oldSB
    Application.DisplayAlerts = oldDA
    Application.EnableEvents = oldEV
    Application.ScreenUpdating = oldSU
    Status_Show
    Exit Sub

Fatal:
    Status_Finish 0, total, 0, "Stopped: " & Err.Number & " : " & Err.Description
    MsgBox "Build stopped: " & Err.Number & " : " & Err.Description, vbCritical, "Second Term Analysis"
    Resume Done
End Sub

' Builds, saves and closes one workbook. Returns "" when the file really exists on disk,
' otherwise the reason (never an empty text for a failure).
Private Function BuildOneFile(ByVal stem As String, ByVal titles As Variant, ByVal sheetIndex As Object, _
                              ByVal outDir As String, ByRef savedPath As String) As String
    Dim wbNew As Workbook, errText As String, acc As String

    On Error GoTo EH
    Application.Calculation = xlCalculationManual
    Set wbNew = Workbooks.Add(xlWBATWorksheet)

    If stem = "1LitAnalysis02" Then
        acc = baf_ArT("645 639 62F 644 20 627 644 642 628 648 644")
    Else
        acc = baf_ArT("645 639 62F 644 20 627 644 627 646 62A 642 627 644")
    End If

    If Not baf_BuildDataSheet(wbNew, sheetIndex, titles, acc, errText) Then Err.Raise vbObjectError + 901, "Build_All", errText
    If Not baf_BuildAnalysisSheets(wbNew, baf_CoreKeysFor(stem), errText) Then Err.Raise vbObjectError + 902, "Build_All", errText
    If Not baf_BuildCompareSheet(wbNew, errText) Then Err.Raise vbObjectError + 905, "Build_All", errText
    If Not baf_BuildCartTecSheet(wbNew, errText) Then Err.Raise vbObjectError + 903, "Build_All", errText

    On Error Resume Next
    wbNew.Worksheets(baf_DATA_SHEET).Activate
    wbNew.Worksheets(baf_DATA_SHEET).Range("A1").Select
    On Error GoTo EH

    ' one full calculation, then the file is stored in automatic mode
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    baf_AutoFitOutput wbNew
    errText = baf_ProtectOutput(wbNew)
    If Len(errText) > 0 Then Err.Raise vbObjectError + 906, "Build_All", errText

    savedPath = baf_NextFreePath(outDir, stem, ".xlsx")
    wbNew.SaveAs Filename:=savedPath, FileFormat:=xlOpenXMLWorkbook
    If Not baf_FileExistsP(savedPath) Then Err.Raise vbObjectError + 904, "Build_All", "The file was not found after saving: " & savedPath

    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing
    BuildOneFile = ""
    Exit Function
EH:
    BuildOneFile = Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
End Function

Private Sub WriteLog(ByVal path As String, ByVal txt As String)
    Dim fso As Object, ts As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, True)
    ts.Write txt
    ts.Close
    Err.Clear
End Sub

'==============================================================================
' STATUS SHEET "Term02Status"
'   Created by BuildControlPanel (button that builds the control panel) through
'   Status_CreateSheet; updated by Build_All while it works.
'==============================================================================
Private Function StatusWs() As Worksheet
    On Error Resume Next
    Set StatusWs = ThisWorkbook.Worksheets(STATUS_SHEET)
    Err.Clear
End Function

Public Sub Status_CreateSheet()
    Dim ws As Worksheet, prev As Object, map As Object, k As Variant, i As Long, n As Long, base As String

    On Error GoTo EH
    baf_EnsureFonts
    Set prev = ActiveSheet

    Set ws = StatusWs()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = STATUS_SHEET
    Else
        ws.Cells.UnMerge
        ws.Cells.Clear
    End If

    ws.DisplayRightToLeft = True
    With ws.Cells
        .Font.Name = baf_gFontBase
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With

    ws.Columns(1).ColumnWidth = 2.5
    ws.Columns(2).ColumnWidth = 5
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 20
    ws.Columns(6).ColumnWidth = 62
    ws.Columns(7).ColumnWidth = 11
    ws.Columns(8).ColumnWidth = 2.5

    ' ---- banner
    ws.Rows(1).RowHeight = 10
    baf_MergeText ws, 2, 2, 7, baf_ArT("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A")
    With ws.Range("B2:G2")
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = baf_gFontBase: .Font.Size = 28: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(2).RowHeight = 58

    baf_MergeText ws, 3, 2, 7, baf_ArT("625 635 62F 627 631 20 645 644 641 627 62A 20 627 644 62A 62D 644 64A 644 20 28 44 61 74 61 20 20 2D 20 20 41 6E 61 6C 79 73 69 73 30 31 20 20 2D 20 20 41 6E 61 6C 79 73 69 73 30 32 20 20 2D 20 20 43 6F 6D 70 61 72 65 20 20 2D 20 20 43 61 72 74 5F 54 65 63 29 20 644 643 644 20 627 644 634 639 628 20 628 636 63A 637 629 20 648 627 62D 62F 629")
    With ws.Range("B3:G3")
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = baf_gFontName: .Font.Size = 14
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 12

    ' ---- information block
    base = ThisWorkbook.Path
    StatusInfo ws, 5, baf_ArT("645 62C 644 62F 20 627 644 645 635 62F 631"), base & "\Source02"
    StatusInfo ws, 6, baf_ArT("645 62C 644 62F 20 627 644 625 635 62F 627 631"), base & "\" & baf_OutputFolderName()
    StatusInfo ws, 7, baf_ArT("622 62E 631 20 62A 634 63A 64A 644"), "-"
    StatusInfo ws, 8, baf_ArT("627 644 646 62A 64A 62C 629"), "-"
    ws.Rows(9).RowHeight = 14

    ' ---- status table
    baf_MergeText ws, 10, 2, 7, baf_ArT("62D 627 644 629 20 627 644 645 644 641 627 62A")
    With ws.Range("B10:G10")
        .Font.Name = baf_gFontBase: .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlRight
        .ReadingOrder = xlRTL
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(28, 55, 103)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(10).RowHeight = 30

    ws.Range("B11").Value2 = "#"
    ws.Range("C11").Value2 = baf_ArT("627 644 645 644 641")
    ws.Range("D11").Value2 = baf_ArT("627 644 623 642 633 627 645")
    ws.Range("E11").Value2 = baf_ArT("627 644 62D 627 644 629")
    ws.Range("F11").Value2 = baf_ArT("627 644 62A 641 627 635 64A 644")
    ws.Range("G11").Value2 = baf_ArT("627 644 645 62F 629 20 28 62B 29")
    baf_StyleHead ws.Range("B11:G11")
    ws.Rows(11).RowHeight = 28

    Set map = baf_GetFileMap()
    n = 0
    For Each k In map.Keys
        n = n + 1
        With ws.Cells(ROW_FIRST + n - 1, 2)
            .Value2 = n
            .NumberFormat = "00"
        End With
        ws.Cells(ROW_FIRST + n - 1, 3).Value2 = CStr(k)
        ws.Cells(ROW_FIRST + n - 1, 4).Value2 = UBound(map(k)) - LBound(map(k)) + 1
    Next k
    With ws.Range(ws.Cells(ROW_FIRST, 2), ws.Cells(ROW_FIRST + n - 1, 7))
        .Font.Name = baf_gFontNum: .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    With ws.Range(ws.Cells(ROW_FIRST, 6), ws.Cells(ROW_FIRST + n - 1, 6))
        .HorizontalAlignment = xlRight
        .WrapText = True
        .Font.Size = 12
    End With
    For i = 1 To n
        Status_Row i, 0, "", 0
    Next i

    ' ---- footer
    ws.Rows(ROW_FIRST + n).RowHeight = 12
    baf_MergeText ws, ROW_FIRST + n + 1, 2, 7, baf_ArT("644 627 20 64A 64F 633 62A 628 62F 644 20 623 64A 20 645 644 641 20 645 648 62C 648 62F 3A 20 639 646 62F 20 625 639 627 62F 629 20 627 644 62A 634 63A 64A 644 20 62A 64F 62D 641 638 20 627 644 645 644 641 627 62A 20 627 644 62C 62F 64A 62F 629 20 628 627 644 644 627 62D 642 629 20 5F 30 31 20 648 20 5F 30 32 20 2E 2E 2E 20 648 64A 64F 62D 641 638 20 633 62C 644 20 643 627 645 644 20 644 643 644 20 62A 634 63A 64A 644 20 641 64A 20 645 62C 644 62F 20 627 644 625 635 62F 627 631 2E")
    With ws.Range(ws.Cells(ROW_FIRST + n + 1, 2), ws.Cells(ROW_FIRST + n + 1, 7))
        .Font.Name = baf_gFontName: .Font.Size = 12
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlRight
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
    ws.Rows(ROW_FIRST + n + 1).RowHeight = 34

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select
    prev.Activate
    Err.Clear
    Exit Sub
EH:
    MsgBox "Status_CreateSheet: " & Err.Number & " : " & Err.Description, vbExclamation, "Second Term Analysis"
End Sub

Private Sub StatusInfo(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal txt As String)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Merge
    ws.Cells(r, 2).Value2 = lbl
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = baf_gFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 7)).Merge
    ws.Cells(r, 4).NumberFormat = "@"
    ws.Cells(r, 4).Value2 = txt
    With ws.Range(ws.Cells(r, 4), ws.Cells(r, 7))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Name = baf_gFontNum: .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With
    ws.Rows(r).RowHeight = 28
End Sub

' kind: 0 waiting, 1 working, 2 done, 3 failed
Private Sub Status_Row(ByVal idx As Long, ByVal kind As Long, ByVal detail As String, ByVal secs As Double)
    Dim ws As Worksheet, r As Long
    Set ws = StatusWs()
    If ws Is Nothing Then Exit Sub
    baf_EnsureFonts
    r = ROW_FIRST + idx - 1
    On Error Resume Next
    With ws.Cells(r, 5)
        Select Case kind
            Case 0
                .Value2 = baf_ArT("641 64A 20 627 644 627 646 62A 638 627 631")
                .Interior.Color = RGB(242, 242, 242): .Font.Color = RGB(100, 100, 100)
            Case 1
                .Value2 = baf_ArT("642 64A 62F 20 627 644 625 646 62C 627 632")
                .Interior.Color = RGB(255, 242, 204): .Font.Color = RGB(156, 101, 0)
            Case 2
                .Value2 = baf_ArT("62A 645")
                .Interior.Color = RGB(198, 239, 206): .Font.Color = RGB(0, 97, 0)
            Case Else
                .Value2 = baf_ArT("641 634 644")
                .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
        End Select
        .Font.Bold = True
        .Font.Name = baf_gFontBase
    End With
    ws.Cells(r, 6).NumberFormat = "@"
    ws.Cells(r, 6).Value2 = detail
    If kind >= 2 Then
        ws.Cells(r, 7).Value2 = Round(secs, 1)
        ws.Cells(r, 7).NumberFormat = "0.0"
    ElseIf kind = 0 Then
        ws.Cells(r, 7).ClearContents
    End If
    If kind = 3 Then
        ws.Rows(r).RowHeight = 54
    Else
        ws.Rows(r).RowHeight = 26
    End If
    Err.Clear
End Sub

Private Sub Status_Reset(ByVal map As Object)
    Dim ws As Worksheet, i As Long
    Set ws = StatusWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To map.Count
        Status_Row i, 0, "", 0
    Next i
    ws.Cells(7, 4).Value2 = Format$(Now, "yyyy-mm-dd  hh:nn")
    ws.Cells(8, 4).Value2 = baf_ArT("642 64A 62F 20 627 644 62A 646 641 64A 630 20 2E 2E 2E")
    ws.Cells(8, 4).Font.Color = RGB(156, 101, 0)
    Err.Clear
End Sub

Private Sub Status_Finish(ByVal okCount As Long, ByVal total As Long, ByVal secs As Double, ByVal note As String)
    Dim ws As Worksheet
    Set ws = StatusWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If note <> "" Then
        ws.Cells(8, 4).Value2 = note
        ws.Cells(8, 4).Font.Color = RGB(156, 0, 6)
    Else
        ws.Cells(8, 4).Value2 = baf_ArT("62A 645 20 625 646 634 627 621 20") & okCount & baf_ArT("20 20 645 646 20 20") & total & baf_ArT("20 20 645 644 641 627 64B 20 20 2D 20 20") & Format$(secs, "0") & baf_ArT("20 20 62B 627 646 64A 629")
        If okCount = total Then
            ws.Cells(8, 4).Font.Color = RGB(0, 97, 0)
        Else
            ws.Cells(8, 4).Font.Color = RGB(156, 0, 6)
        End If
    End If
    ws.Cells(8, 4).Font.Bold = True
    Err.Clear
End Sub

Private Sub Status_Show()
    Dim ws As Worksheet
    Set ws = StatusWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
    ws.Range("A1").Select
    Err.Clear
End Sub
