Attribute VB_Name = "modOrientationAll02"
Option Explicit

'==============================================================================
' Module  : modOrientationAll02
' Purpose : ONE entry (Run_OrientationAll02) for the whole school-orientation block
'           of Term 02. It is the action of the button "Term 02 Orientation".
'
'   Run_OrientationAll02
'    0. checks that every file needed by the four steps exists, in the order of the steps
'    1. modTransfer02            Data Transfer       Source02 -> Term02Guidance.xlsx
'    2. modStreamPreferences02   Stream Preferences  StreamPreferences02.xlsx -> wishes
'    3. modOrientation02         Orientation         decisions + general report
'    4. modOrientationResults02  Export results      -> \تحليل نتائج الفصل الثاني\OrientationResults02[_NN].xlsx
'
' Runs all steps silently and shows ONE consolidated message at the end
' (or ONE diagnostic message if a step stops).
' Handles user interruption (Esc / Ctrl+Break) cleanly so the button never freezes.
'==============================================================================

Private Const TITLE_TXT As String = "Term 02 Orientation"
Private Const SRC_FOLDER As String = "Source02"
Private Const GUIDANCE_FILE As String = "Term02Guidance.xlsx"
Private Const WISHES_FILE As String = "StreamPreferences02.xlsx"

Private m_busy As Boolean

Public Sub ResetOrientationState()
    m_busy = False
    On Error Resume Next
    modTransfer02.SetSilentMode False
    modStreamPreferences02.SetSilentMode False
    modOrientation02.SetSilentMode False
    modOrientationResults02.SetSilentMode False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo 0
End Sub

Public Sub Run_OrientationAll02()
    Dim base As String, problem As String, t0 As Date, outPath As String, secs As Long
    Dim finalMsg As String

    ' Always clear any stale lock from a previously interrupted session
    m_busy = False

    base = ThisWorkbook.Path
    If Len(base) = 0 Then
        MsgBox "Save the tool workbook first: the files are looked for next to it.", vbExclamation, TITLE_TXT
        Exit Sub
    End If
    If InStr(1, base, "://") > 0 Then
        MsgBox "The tool is opened from a cloud/web location:" & vbCrLf & base & vbCrLf & vbCrLf & _
               "Run it from a normal local folder that contains the files.", vbExclamation, TITLE_TXT
        Exit Sub
    End If
    If Right$(base, 1) <> Application.PathSeparator Then base = base & Application.PathSeparator

    ' ---- 0. every file the four steps need, in the order of the steps
    problem = CheckFiles(base)
    If Len(problem) > 0 Then
        MsgBox problem, vbExclamation, TITLE_TXT
        Exit Sub
    End If

    m_busy = True
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    t0 = Now

    ' ---- 1. data transfer (silent)
    modTransfer02.SetSilentMode True
    modTransfer02.Run_Transfer02
    If Not modTransfer02.Transfer02Succeeded() Then
        StopAt 1, "Data Transfer", modTransfer02.LastReport()
        GoTo Finished
    End If

    ' ---- 2. stream preferences / wishes (silent)
    modStreamPreferences02.SetSilentMode True
    modStreamPreferences02.Run_StreamPreferences02
    If Not modStreamPreferences02.Preferences02Succeeded() Then
        StopAt 2, "Stream Preferences", modStreamPreferences02.LastReport()
        GoTo Finished
    End If

    ' ---- 3. orientation (silent)
    modOrientation02.SetSilentMode True
    modOrientation02.Run_Orientation02
    If Not modOrientation02.Orientation02Succeeded() Then
        StopAt 3, "Orientation", modOrientation02.LastReport()
        GoTo Finished
    End If

    ' ---- 4. export results (silent)
    modOrientationResults02.SetSilentMode True
    modOrientationResults02.Run_OrientationResults02
    outPath = modOrientationResults02.Results02OutputPath()
    If Not modOrientationResults02.OrientationResults02Succeeded() Or Len(outPath) = 0 Or Not FileThere(outPath) Then
        StopAt 4, "Export Orientation Results", modOrientationResults02.LastReport()
        GoTo Finished
    End If

    secs = CLng((Now - t0) * 86400)
    finalMsg = "All 4 steps completed (" & secs & " s):" & vbCrLf & vbCrLf & _
               "[1] Data Transfer:" & vbCrLf & modTransfer02.LastReport() & vbCrLf & vbCrLf & _
               "[2] Stream Preferences:" & vbCrLf & modStreamPreferences02.LastReport() & vbCrLf & vbCrLf & _
               "[3] Orientation:" & vbCrLf & modOrientation02.LastReport() & vbCrLf & vbCrLf & _
               "[4] Export Results:" & vbCrLf & modOrientationResults02.LastReport()

    MsgBox finalMsg, vbInformation, TITLE_TXT

Finished:
    On Error Resume Next
    m_busy = False
    modTransfer02.SetSilentMode False
    modStreamPreferences02.SetSilentMode False
    modOrientation02.SetSilentMode False
    modOrientationResults02.SetSilentMode False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo 0
    Exit Sub

Failed:
    If Err.Number = 18 Then
        MsgBox "Execution was stopped by the user. Ready to run again.", vbExclamation, TITLE_TXT
    Else
        MsgBox "The orientation sequence stopped: " & Err.Number & " : " & Err.Description, vbCritical, TITLE_TXT
    End If
    Resume Finished
End Sub

Private Sub StopAt(ByVal stepNo As Long, ByVal stepName As String, ByVal stepReport As String)
    Dim msg As String
    msg = "The sequence stopped at step " & stepNo & " (" & stepName & ")." & vbCrLf & _
          "The steps after it were not run."
    If Len(Trim$(stepReport)) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Details:" & vbCrLf & stepReport
    End If
    MsgBox msg, vbExclamation, TITLE_TXT
End Sub

' Returns "" when everything is there, otherwise the text for the FIRST missing item.
Private Function CheckFiles(ByVal base As String) As String
    If Not FolderHasExcelFiles(base & SRC_FOLDER) Then
        CheckFiles = "Step 1 (Data Transfer) needs the folder '" & SRC_FOLDER & "' with at least one Excel file in it:" & _
                     vbCrLf & base & SRC_FOLDER & vbCrLf & vbCrLf & "Put the term result files there, then run again."
        Exit Function
    End If
    If Not FileThere(base & GUIDANCE_FILE) Then
        CheckFiles = "Step 1 (Data Transfer) needs the file '" & GUIDANCE_FILE & "' next to the tool:" & _
                     vbCrLf & base & vbCrLf & vbCrLf & "Provide that file, then run again."
        Exit Function
    End If
    If Not FileThere(base & WISHES_FILE) Then
        CheckFiles = "Step 2 (Stream Preferences) needs the file '" & WISHES_FILE & "' next to the tool:" & _
                     vbCrLf & base & vbCrLf & vbCrLf & "Provide that file, then run again. Nothing has been changed."
        Exit Function
    End If
End Function

Private Function FileThere(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileThere = fso.FileExists(p)
    If Err.Number <> 0 Then FileThere = False: Err.Clear
End Function

' At least one .xls / .xlsx / .xlsm / .xlsb file that is not an Excel lock file (~$...)
Private Function FolderHasExcelFiles(ByVal p As String) As Boolean
    Dim fso As Object, f As Object, ext As String
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then Exit Function
    For Each f In fso.GetFolder(p).Files
        If Left$(f.Name, 2) <> "~$" Then
            ext = LCase$(fso.GetExtensionName(f.Name))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Then
                FolderHasExcelFiles = True
                Exit Function
            End If
        End If
    Next f
Done:
    Err.Clear
End Function