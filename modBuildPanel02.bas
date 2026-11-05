Attribute VB_Name = "modBuildPanel02"
Option Explicit

' Import only the matching term panel into each independent tool workbook.
Private Const PANEL_SHEET As String = "ControlPanel"
Private Const RESULTS_FILE As String = "Term02Results.xlsx"
Private Const SOURCE_FOLDER As String = "Source02"
Private Const BUTTON_WIDTH As Double = 164
Private Const BUTTON_HEIGHT As Double = 54
Private Const BUTTON_GAP As Double = 12

Private Function PanelActions() As Variant
    ' Second Term Analysis, Cumulative Term Average and Split Term Results work on the raw files of Source02: no previous step is required.
    PanelActions = Array( _
        Array("btnCollect", "Second Term Results", "CollectTerm02Results", False, &H60AE27), _
        Array("btnDistribute", "Second Term Analysis", "modBiuldAnalysisfilesBuildAll02.Build_All", False, &HB98029), _
        Array("btnAverage", "Cumulative Term Average", "CumulativeTermAverage", False, &HAD448E), _
        Array("btnSplitTerms", "Split Term Results", "modSplitTermResults.SplitTermResults", False, &H85A016), _
        Array("btnExamen", "Examen Analysis", "modExamenAnalysis02.Run_ExamenAnalysis02", False, &HA05568), _
        Array("btnCompare", "Compare", "modCompare.Run_Compare", False, &H00B0E8F0), _
        Array("btnOrientationAll", "Term 02 Orientation", "modOrientationAll02.Run_OrientationAll02", False, &HB98029))
End Function

Public Sub BuildControlPanel()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, count As Long, columns As Long, gridRow As Long, gridCol As Long
    Dim rowItems As Long, canvasWidth As Double, rowWidth As Double
    Dim originLeft As Double, originTop As Double, x As Double, y As Double
    Dim prevScreen As Boolean, errorText As String

    prevScreen = Application.ScreenUpdating
    On Error GoTo BuildError
    Application.ScreenUpdating = False
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 801, , "The tool workbook is read-only."
    Set ws = PanelSheet()
    If ws Is Nothing Then
        If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 802, , "The workbook structure is protected."
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET
    End If
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        Err.Raise vbObjectError + 803, , "Unprotect the control panel sheet before rebuilding it."
    End If
    ws.Activate
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    With ws
        .Cells.UnMerge
        .Cells.Clear
        .DisplayRightToLeft = True
        .Columns("A").ColumnWidth = 3
        .Columns("B:K").ColumnWidth = 9.6
        .Rows(1).RowHeight = 48
        .Rows(2).RowHeight = 14
        With .Range("B1:K1")
            .Merge
            .Value = "Second Term Management Dashboard"
            .Font.Name = "Segoe UI"
            .Font.Size = 19
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(28, 55, 103)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End With

    actions = PanelActions()
    count = UBound(actions) - LBound(actions) + 1
    columns = 3
    canvasWidth = columns * BUTTON_WIDTH + (columns - 1) * BUTTON_GAP
    originLeft = ws.Range("B3").Left + (ws.Range("B1:K1").Width - canvasWidth) / 2
    originTop = ws.Range("B3").Top
    For i = 0 To count - 1
        item = actions(i)
        gridRow = i \ columns
        gridCol = i Mod columns
        rowItems = count - gridRow * columns
        If rowItems > columns Then rowItems = columns
        rowWidth = rowItems * BUTTON_WIDTH + (rowItems - 1) * BUTTON_GAP
        x = originLeft + (canvasWidth - rowWidth) / 2 + gridCol * (BUTTON_WIDTH + BUTTON_GAP)
        y = originTop + gridRow * (BUTTON_HEIGHT + BUTTON_GAP)
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, BUTTON_WIDTH, BUTTON_HEIGHT)
        With btn
            .Name = CStr(item(0))
            .Adjustments(1) = 0.12
            .Placement = xlFreeFloating
            .Line.Weight = 1.25
            With .TextFrame2
                .TextRange.Text = CStr(item(1))
                .TextRange.Font.Name = "Segoe UI"
                .TextRange.Font.Size = 12
                .TextRange.Font.Bold = msoTrue
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .VerticalAnchor = msoAnchorMiddle
                .WordWrap = msoTrue
                .MarginLeft = 7
                .MarginRight = 7
                .MarginTop = 4
                .MarginBottom = 4
            End With
        End With
    Next i
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    RefreshPanelState
    CheckPanelMacros actions
    RemoveOldWorkflowNames
    BuildStatusSheetSafe
    ws.Activate
    Application.ScreenUpdating = prevScreen
    MsgBox "Control panel updated: " & count & " buttons in three columns.", vbInformation, "Done"
    Exit Sub
BuildError:
    errorText = CStr(Err.Number) & " : " & Err.Description
    Application.ScreenUpdating = prevScreen
    MsgBox "Panel build failed: " & errorText, vbExclamation, "Error"
End Sub

Public Sub RefreshPanelState()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, enabled As Boolean, resultsExist As Boolean
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    resultsExist = PanelResultsExist()
    actions = PanelActions()
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        Set btn = Nothing
        On Error Resume Next
        Set btn = ws.Shapes(CStr(item(0)))
        Err.Clear
        On Error GoTo 0
        If Not btn Is Nothing Then
            enabled = True
            If CBool(item(3)) Then enabled = resultsExist
            With btn
                .Fill.Solid
                If enabled Then
                    .Fill.ForeColor.RGB = CLng(item(4))
                    .Line.ForeColor.RGB = CLng(item(4))
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .OnAction = PanelMacroAction(CStr(item(2)))
                Else
                    .Fill.ForeColor.RGB = RGB(220, 220, 220)
                    .Line.ForeColor.RGB = RGB(180, 180, 180)
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
                    .OnAction = PanelMacroAction("ResultsGuard")
                End If
            End With
        End If
    Next i
End Sub

Public Sub ResultsGuard()
    MsgBox "Run Second Term Results first to generate " & RESULTS_FILE & ".", _
           vbExclamation, "Step 1 Required"
End Sub

' Compatibility with buttons created by earlier panel versions.
Public Sub DistributeGuard()
    ResultsGuard
End Sub
Public Sub AverageGuard()
    ResultsGuard
End Sub

' Deletes the hidden names (_Term02_*) written by the old step-by-step workflow, which no longer exists.
Private Sub RemoveOldWorkflowNames()
    Dim i As Long
    On Error Resume Next
    For i = ThisWorkbook.Names.Count To 1 Step -1
        If InStr(1, ThisWorkbook.Names(i).Name, "_Term02_", vbTextCompare) > 0 Then ThisWorkbook.Names(i).Delete
    Next i
    Err.Clear
End Sub

' Creates / refreshes the status sheet of Second Term Analysis (source folder, output folder,
' last run, duration, success and the time of every file). Silent when the module is not imported.
Private Sub BuildStatusSheetSafe()
    On Error Resume Next
    Application.Run PanelMacroAction("modBiuldAnalysisfilesBuildAll02.Status_CreateSheet")
    Err.Clear
End Sub

Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

Private Function PanelResultsExist() As Boolean
    Dim attrs As Long, p As String
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Function
    p = ThisWorkbook.Path & Application.PathSeparator & PanelOutputFolderName() & Application.PathSeparator & RESULTS_FILE
    On Error GoTo Missing
    attrs = GetAttr(p)
    PanelResultsExist = ((attrs And vbDirectory) = 0)
    Exit Function
Missing:
    PanelResultsExist = False
End Function

' Output folder that holds Term02Results.xlsx (hex code points, same name as in modCollector)
Private Function PanelOutputFolderName() As String
    Dim parts() As String, i As Long, s As String
    parts = Split("0645 0639 0627 0644 062C 0629 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0641 0635 0644 0020 0030 0032", " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    PanelOutputFolderName = s
End Function

Private Function PanelMacroAction(ByVal macroName As String) As String
    PanelMacroAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

Private Sub CheckPanelMacros(ByVal actions As Variant)
    Dim i As Long, item As Variant, missing As String
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        If Not PanelMacroExists(CStr(item(2))) Then missing = missing & "  - " & CStr(item(2)) & vbCrLf
    Next i
    If Len(missing) > 0 Then
        MsgBox "Import the missing modules before using these actions:" & vbCrLf & missing, _
               vbExclamation, "Missing Macros"
    End If
End Sub

Private Function PanelMacroExists(ByVal macroName As String) As Boolean
    Dim comp As Object, codeMod As Object, txt As String
    Dim moduleName As String, procedureName As String, dotPos As Long
    On Error GoTo AccessDenied
    dotPos = InStrRev(macroName, ".")
    procedureName = macroName
    If dotPos > 0 Then
        moduleName = Left$(macroName, dotPos - 1)
        procedureName = Mid$(macroName, dotPos + 1)
    End If
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then
            If Len(moduleName) = 0 Or StrComp(comp.Name, moduleName, vbTextCompare) = 0 Then
                Set codeMod = comp.CodeModule
                If codeMod.CountOfLines > 0 Then
                    txt = codeMod.Lines(1, codeMod.CountOfLines)
                    If InStr(1, txt, "Sub " & procedureName & "(", vbTextCompare) > 0 Then
                        PanelMacroExists = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next comp
    Exit Function
AccessDenied:
    ' Macro inspection is optional when VBA project access is disabled.
    PanelMacroExists = True
    Err.Clear
End Function
