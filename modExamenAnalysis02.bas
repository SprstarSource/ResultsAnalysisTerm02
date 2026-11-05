Attribute VB_Name = "modExamenAnalysis02"
Option Explicit

'==============================================================================
' modExamenAnalysis02  -  unified module (second semester)
'   1) reads every .xlsx in the folder  ExamenSource02  (next to this workbook)
'   2) consolidates marks per class  (Philo, LEtr, SE01, SE02, Math, GEc)
'   3) builds a NEW workbook:  Gen + 6 class sheets + data-transfer summary
'      - semester average computed here (coefficients per class)
'      - class tables sorted by average (descending), analysis tables below
'      - analysis cells are fixed values (no formulas) except the live cells of Gen:
'        the "difference" column of section 3 and two hidden helper columns (H:I)
'        of section 4, which drive the conditional formatting (manual input)
'   4) saves it in the output folder next to this workbook (created when missing):
'      <Arabic name> .xlsx / _01 / _02 ...
' Subject recognition: exact keywords first; if none matches, a one-wrong-letter
'   fallback (unique subject only) - every such guess is written in the transfer sheet.
' All Arabic text is stored as hex code points (ArHex) so this file is pure ASCII.
'==============================================================================

' ---------------- configuration ----------------
Private Const SOURCE_FOLDER_NAME As String = "ExamenSource02"
Private Const SRC_HEADER_ROW As Long = 8
Private Const DATA_START_ROW As Long = 9
Private Const MAX_STUDENTS As Long = 200
Private Const MARK_OFFSET As Long = 0
Private Const ABSENT_IF_NO_TEST_COLUMN As Boolean = False
Private Const BLOCK_ON_ROSTER_MISMATCH As Boolean = True
Private Const PASS_MARK As Double = 10
Private Const SEASON_START_MONTH As Long = 9

' ---------------- layout (same values as modBacEtab / modBacAnalysis) ----------------
Private Const TITLE_ROW As Long = 1
Private Const HEADER_ROW As Long = 2
Private Const FIRST_DATA_ROW As Long = 3
Private Const FIRST_NUM_COL As Long = 4
Private Const GAP_BELOW_DATA As Long = 2
Private Const TITLE_ROW_HEIGHT As Double = 40
Private Const HEADER_ROW_HEIGHT As Double = 100
Private Const DATA_ROW_HEIGHT As Double = 24
Private Const HDR_ROW_HEIGHT As Double = 100
Private Const TITLE_HEIGHT As Double = 36
Private Const GEN_HEAD_HEIGHT As Double = 40
Private Const ID_COL_WIDTH As Double = 4.3
Private Const NUM_COL_WIDTH As Double = 6.3
Private Const DATE_COL_MIN_WIDTH As Double = 10.5

Private Const FMT_COUNT As String = "00"
Private Const FMT_DEC As String = "00.00"
Private Const FMT_PCT As String = "[=1]%0;%00.00"

Private Const CHECK_FONTS As Boolean = True
Private Const FONT_BASE As String = "Dubai"
Private Const FONT_NAME As String = "Dubai Medium"
Private Const FONT_NUM As String = "Aptos Slab"
Private Const FONT_SAFE_TEXT As String = "Arial"
Private Const FONT_SAFE_NUM As String = "Calibri"

' ---------------- types / state ----------------
Private Type ColStat
    N As Long
    Mean As Double
    Median As Double
    ModeV As Variant
    MinV As Double
    MaxV As Double
    Sd As Double
    Vari As Double
    Pass As Long
    Bands(0 To 7) As Long
    Cls(1 To 5) As Long
End Type

Private mFontBase As String
Private mFontName As String
Private mFontNum As String
Private mFontNotes As String
Private mLayoutNotes As String

Private g_wbSource As Workbook
Private g_wbOut As Workbook
Private g_OpenFailures As String
Private g_ModelFailures As String
Private g_CountConflicts As String
Private g_RosterConflicts As String
Private g_Notes As String
Private g_Fuzzy As String
Private g_CohortCount As Object
Private g_CohortDisplay As Object
Private g_CohortOrigin As Object
Private g_SubjectDisplay As Object
Private g_Seen As Object

Private mExemptKey As String
Private mN(1 To 6) As Long
Private mRoster(1 To 6) As String
Private mReg(1 To 6) As Long
Private mPres(1 To 6) As Long
Private mPass(1 To 6) As Long
Private mExempt(1 To 6) As Long
Private mNames(1 To 6, 1 To MAX_STUDENTS) As Variant
Private mBirth(1 To 6, 1 To MAX_STUDENTS) As Variant
Private mMarks(1 To 6, 1 To MAX_STUDENTS, 1 To 12) As Variant
Private mLoaded(1 To 6, 1 To 12) As Boolean

'==============================================================================
'  Entry point
'==============================================================================
Public Sub Run_ExamenAnalysis02()
    Dim baseFolder As String, sourceFolder As String, outDir As String, outputPath As String
    Dim dictSummary As Object
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean, oldStatus As Variant
    Dim stateSaved As Boolean, Failed As Boolean, savedOK As Boolean
    Dim errText As String, msg As String, s As Long, sheetsMade As Long

    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatus = Application.StatusBar
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    InitState

    baseFolder = ThisWorkbook.Path
    If baseFolder = "" Or LCase$(Left$(baseFolder, 4)) = "http" Then
        Err.Raise vbObjectError + 901, "modExamenAnalysis02", Ar("save_tool")
    End If
    sourceFolder = baseFolder & Application.PathSeparator & SOURCE_FOLDER_NAME & Application.PathSeparator
    If Not FolderExists(sourceFolder) Then
        Err.Raise vbObjectError + 902, "modExamenAnalysis02", Ar("no_folder") & vbLf & sourceFolder
    End If

    Set dictSummary = CreateObject("Scripting.Dictionary")
    ConsolidateSources sourceFolder, dictSummary
    CheckCoverage

    Application.StatusBar = Ar("working") & " ..."
    BuildOutputWorkbook dictSummary

    outDir = baseFolder & Application.PathSeparator & OutputFolderName()
    If Not EnsureFolder(outDir) Then Err.Raise vbObjectError + 904, "modExamenAnalysis02", "The output folder could not be created: " & outDir
    outputPath = NewOutputPath(outDir, Ar("out_name"))
    g_wbOut.CheckCompatibility = False
    g_wbOut.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    savedOK = True
    sheetsMade = g_wbOut.Worksheets.Count

Finished:
    On Error Resume Next
    If Not g_wbSource Is Nothing Then g_wbSource.Close SaveChanges:=False
    Set g_wbSource = Nothing
    If Failed And Not savedOK And Not g_wbOut Is Nothing Then g_wbOut.Close SaveChanges:=False
    If stateSaved Then
        Application.StatusBar = oldStatus
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    On Error GoTo 0

    If Failed Then
        MsgBox Ar("failed") & vbLf & errText, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("app_title2")
    Else
        msg = Ar("done") & vbLf & outputPath & vbLf & vbLf & Ar("sheets") & " : " & sheetsMade
        For s = 1 To 6
            msg = msg & vbLf & SheetName(s) & " : " & mReg(s) & " / " & mPres(s)
        Next s
        If mFontNotes <> "" Then msg = msg & vbLf & vbLf & Ar("fonts") & mFontNotes
        If mLayoutNotes <> "" Then msg = msg & vbLf & vbLf & Ar("layout_warn") & mLayoutNotes
        If HasProblems() Then
            msg = msg & vbLf & vbLf & Ar("see_sheet")
        Else
            msg = msg & vbLf & vbLf & Ar("ok_all")
        End If
        If Len(msg) > 1000 Then msg = Left$(msg, 1000) & vbLf & "..."
        MsgBox msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, Ar("app_title2")
    End If
    Exit Sub

Failed:
    Failed = True
    errText = Err.Description
    Resume Finished
End Sub

Private Function HasProblems() As Boolean
    HasProblems = (g_OpenFailures <> "" Or g_ModelFailures <> "" Or g_CountConflicts <> "" _
                   Or g_RosterConflicts <> "" Or g_Notes <> "" Or g_Fuzzy <> "")
End Function

Private Sub InitState()
    Dim s As Long
    g_OpenFailures = "": g_ModelFailures = "": g_CountConflicts = "": g_RosterConflicts = "": g_Notes = "": g_Fuzzy = ""
    mFontNotes = "": mLayoutNotes = ""
    Set g_CohortCount = CreateObject("Scripting.Dictionary")
    Set g_CohortDisplay = CreateObject("Scripting.Dictionary")
    Set g_CohortOrigin = CreateObject("Scripting.Dictionary")
    Set g_SubjectDisplay = CreateObject("Scripting.Dictionary")
    Set g_Seen = CreateObject("Scripting.Dictionary")
    Set g_wbSource = Nothing
    Set g_wbOut = Nothing
    Erase mNames
    Erase mBirth
    Erase mMarks
    Erase mLoaded
    For s = 1 To 6
        mN(s) = 0: mRoster(s) = ""
        mReg(s) = 0: mPres(s) = 0: mPass(s) = 0: mExempt(s) = 0
    Next s
    mExemptKey = SmartClean(Ar("exempt"))
End Sub

'==============================================================================
'  Class / subject definitions (order, coefficients, headers)
'  subject ids: 1 Arabic, 2 Philosophy, 3 Hist-Geo, 4 Islamic, 5 French, 6 English,
'  7 3rd language, 8 Math, 9 Physics, 10 Natural sciences, 11 Accounting,
'  12 Economy, 13 Law, 14 Physical education
'==============================================================================
Private Function SheetName(ByVal s As Long) As String
    SheetName = Choose(s, "Philo", "LEtr", "SE01", "SE02", "Math", "GEc")
End Function

Private Function SheetIds(ByVal s As Long) As Variant
    Select Case s
        Case 1: SheetIds = Array(1, 2, 3, 4, 5, 6, 8, 14)
        Case 2: SheetIds = Array(1, 2, 3, 4, 5, 6, 7, 8, 14)
        Case 3, 4, 5: SheetIds = Array(1, 8, 9, 10, 4, 3, 5, 6, 2, 14)
        Case 6: SheetIds = Array(1, 8, 11, 12, 13, 4, 3, 5, 6, 2, 14)
    End Select
End Function

Private Function SheetCoefs(ByVal s As Long) As Variant
    Select Case s
        Case 1: SheetCoefs = Array(6, 6, 4, 2, 3, 3, 2, 1)
        Case 2: SheetCoefs = Array(5, 2, 2, 2, 5, 5, 4, 2, 1)
        Case 3, 4: SheetCoefs = Array(3, 5, 5, 6, 2, 2, 2, 2, 2, 1)
        Case 5: SheetCoefs = Array(3, 7, 6, 2, 2, 2, 2, 2, 2, 1)
        Case 6: SheetCoefs = Array(3, 5, 6, 5, 2, 2, 4, 2, 2, 2, 1)
    End Select
End Function

Private Function SubjHeader(ByVal id As Long) As String
    SubjHeader = Ar("sj" & CStr(id))
End Function

Private Function SeriesBranch(ByVal s As Long) As String
    SeriesBranch = Ar("bn" & CStr(s))
End Function

Private Function GenBranch(ByVal s As Long) As String
    GenBranch = Ar("gb" & CStr(s))
End Function

' Subject id from a normalized subject key.  -1 = ignored subject, 0 = unknown
Private Function SubjectIdOf(ByVal key As String) As Long
    If InStr(key, Ar("k_amaz")) > 0 Or InStr(key, Ar("k_basic")) > 0 Then SubjectIdOf = -1: Exit Function
    If key = Ar("k_pe") Then SubjectIdOf = 14: Exit Function
    If InStr(key, Ar("k_l3")) > 0 Or InStr(key, Ar("k_ger")) > 0 Or InStr(key, Ar("k_spa")) > 0 _
       Or InStr(key, Ar("k_ita")) > 0 Then SubjectIdOf = 7: Exit Function
    If InStr(key, Ar("k_ar")) > 0 Then SubjectIdOf = 1: Exit Function
    If InStr(key, Ar("k_philo")) > 0 Then SubjectIdOf = 2: Exit Function
    If InStr(key, Ar("k_hist")) > 0 Or InStr(key, Ar("k_geo")) > 0 Then SubjectIdOf = 3: Exit Function
    If InStr(key, Ar("k_isl")) > 0 Then SubjectIdOf = 4: Exit Function
    If InStr(key, Ar("k_fr")) > 0 Then SubjectIdOf = 5: Exit Function
    If InStr(key, Ar("k_en1")) > 0 Or InStr(key, Ar("k_en2")) > 0 Then SubjectIdOf = 6: Exit Function
    If InStr(key, Ar("k_phys")) > 0 Then SubjectIdOf = 9: Exit Function
    If InStr(key, Ar("k_svt")) > 0 Then SubjectIdOf = 10: Exit Function
    If InStr(key, Ar("k_acc")) > 0 Then SubjectIdOf = 11: Exit Function
    If InStr(key, Ar("k_eco")) > 0 Then SubjectIdOf = 12: Exit Function
    If InStr(key, Ar("k_law")) > 0 Then SubjectIdOf = 13: Exit Function
    If InStr(key, Ar("k_math")) > 0 Then SubjectIdOf = 8: Exit Function
End Function

' One-wrong-letter fallback, used only when no exact keyword matched.
' Rules: keyword >= 4 letters (Arabic root excluded: it also matches "tarbiya"),
' substitution only, and the guess is accepted only if exactly one subject matches.
' Physical-education root (3 letters) is allowed, wrong letter only in position 2 or 3.
Private Function FuzzyHit(ByVal key As String, ByVal kw As String, ByVal firstFixed As Boolean) As Boolean
    Dim i As Long, k As Long, bad As Long, L As Long
    L = Len(kw)
    If L = 0 Or Len(key) < L Then Exit Function
    For i = 1 To Len(key) - L + 1
        bad = 0
        For k = 1 To L
            If Mid$(key, i + k - 1, 1) <> Mid$(kw, k, 1) Then
                bad = bad + 1
                If k = 1 And firstFixed Then bad = 2
                If bad > 1 Then Exit For
            End If
        Next k
        If bad <= 1 Then FuzzyHit = True: Exit Function
    Next i
End Function

Private Sub FuzzyProbe(ByVal key As String, ByVal kw As String, ByVal id As Long, _
                       ByRef cand As Long, ByRef multi As Boolean, ByVal firstFixed As Boolean)
    If Len(kw) < 4 And Not firstFixed Then Exit Sub
    If FuzzyHit(key, kw, firstFixed) Then
        If cand = 0 Then
            cand = id
        ElseIf cand <> id Then
            multi = True
        End If
    End If
End Sub

Private Function SubjectIdFuzzy(ByVal key As String) As Long
    Dim cand As Long, multi As Boolean
    FuzzyProbe key, Ar("k_pe_root"), 14, cand, multi, True
    FuzzyProbe key, Ar("k_l3"), 7, cand, multi, False
    FuzzyProbe key, Ar("k_ger"), 7, cand, multi, False
    FuzzyProbe key, Ar("k_spa"), 7, cand, multi, False
    FuzzyProbe key, Ar("k_ita"), 7, cand, multi, False
    FuzzyProbe key, Ar("k_philo"), 2, cand, multi, False
    FuzzyProbe key, Ar("k_hist"), 3, cand, multi, False
    FuzzyProbe key, Ar("k_geo"), 3, cand, multi, False
    FuzzyProbe key, Ar("k_isl"), 4, cand, multi, False
    FuzzyProbe key, Ar("k_fr"), 5, cand, multi, False
    FuzzyProbe key, Ar("k_en1"), 6, cand, multi, False
    FuzzyProbe key, Ar("k_en2"), 6, cand, multi, False
    FuzzyProbe key, Ar("k_phys"), 9, cand, multi, False
    FuzzyProbe key, Ar("k_svt"), 10, cand, multi, False
    FuzzyProbe key, Ar("k_acc"), 11, cand, multi, False
    FuzzyProbe key, Ar("k_eco"), 12, cand, multi, False
    FuzzyProbe key, Ar("k_law"), 13, cand, multi, False
    FuzzyProbe key, Ar("k_math"), 8, cand, multi, False
    If Not multi Then SubjectIdFuzzy = cand
End Function

'==============================================================================
'  PART 1 : read the source folder
'==============================================================================
Private Sub ConsolidateSources(ByVal sourceFolder As String, ByVal dictSummary As Object)
    Dim fileName As String
    Dim files As Collection
    Dim wsSource As Worksheet
    Dim i As Long

    Set files = New Collection
    fileName = Dir(sourceFolder & "*.xlsx")
    Do While fileName <> ""
        If Left$(fileName, 2) <> "~$" Then files.Add fileName
        fileName = Dir()
    Loop

    If files.Count = 0 Then
        Err.Raise vbObjectError + 903, "modExamenAnalysis02", Ar("no_files") & vbLf & sourceFolder
    End If

    For i = 1 To files.Count
        fileName = CStr(files(i))
        Application.StatusBar = Ar("working") & " (" & i & "/" & files.Count & ") " & fileName

        If IsWorkbookOpen(sourceFolder & fileName) Then
            g_OpenFailures = g_OpenFailures & "  >> " & fileName & " " & Ar("fl_already") & vbCrLf
        Else
            Set g_wbSource = OpenSourceSafe(sourceFolder & fileName)
            If g_wbSource Is Nothing Then
                g_OpenFailures = g_OpenFailures & "  >> " & fileName & vbCrLf
            Else
                For Each wsSource In g_wbSource.Worksheets
                    If wsSource.Name <> "Worksheet" Then
                        ProcessSheet wsSource, dictSummary, fileName
                    End If
                Next wsSource
                g_wbSource.Close SaveChanges:=False
                Set g_wbSource = Nothing
            End If
        End If
    Next i
End Sub

Private Function OpenSourceSafe(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(fullPath, ReadOnly:=True, UpdateLinks:=False, AddToMRU:=False)
    If Err.Number <> 0 Then
        Err.Clear
        Set wb = Nothing
    End If
    On Error GoTo 0
    Set OpenSourceSafe = wb
End Function

Private Function IsWorkbookOpen(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Sub LogModelFailure(ByVal label As String, ByVal reason As String)
    g_ModelFailures = g_ModelFailures & "  >> " & label
    If reason <> "" Then g_ModelFailures = g_ModelFailures & " (" & reason & ")"
    g_ModelFailures = g_ModelFailures & vbCrLf
End Sub

Private Sub ProcessSheet(ByVal wsSource As Worksheet, ByVal dictSummary As Object, ByVal fileName As String)
    Dim label As String
    Dim cohortName As String, subjectName As String
    Dim a5 As Variant, textA5 As String
    Dim p1 As Long, p2 As Long, pEnd As Long, colon1 As Long, colon2 As Long
    Dim lastRow As Long, testCol As Long, c As Long
    Dim h As Variant
    Dim modelValid As Boolean
    Dim studentCount As Long
    Dim cohortKey As String, subjKey As String, summaryKey As String, seenKey As String
    Dim s As Long, subjId As Long, fuzzyUsed As Boolean

    label = fileName & " / " & wsSource.Name

    ' ---------- parse A5 ----------
    a5 = wsSource.Range("A5").Value
    If IsError(a5) Then textA5 = "" Else textA5 = CStr(a5)

    modelValid = False
    p1 = InStr(1, textA5, Ar("kw_cohort"))
    If p1 > 0 Then
        colon1 = InStr(p1, textA5, ":")
        If colon1 > 0 Then
            p2 = InStr(colon1 + 1, textA5, Ar("kw_subject"))
            If p2 > 0 Then
                colon2 = InStr(p2, textA5, ":")
                If colon2 > 0 Then
                    pEnd = p2
                    If p2 - 2 > colon1 Then
                        If Mid$(textA5, p2 - 2, 2) = Ar("kw_al") Then pEnd = p2 - 2
                    End If
                    cohortName = CleanLabel(Mid$(textA5, colon1 + 1, pEnd - colon1 - 1))
                    subjectName = CleanLabel(Mid$(textA5, colon2 + 1))
                    modelValid = (cohortName <> "" And subjectName <> "")
                End If
            End If
        End If
    End If

    If Not modelValid Then
        LogModelFailure label, Ar("r_a5")
        Exit Sub
    End If

    ' ---------- subjects that are never transferred (Amazigh / basic subjects) ----------
    subjKey = SmartCleanSubject(subjectName)
    subjId = SubjectIdOf(subjKey)
    If subjId = 0 Then
        subjId = SubjectIdFuzzy(subjKey)
        fuzzyUsed = (subjId > 0)
    End If
    If subjId = -1 Then Exit Sub

    ' ---------- locate the test column (G or H) ----------
    testCol = 0
    For c = 7 To 8
        h = wsSource.Cells(SRC_HEADER_ROW, c).Value
        If Not IsError(h) Then
            If IsTestHeader(CStr(h)) Then
                testCol = c
                Exit For
            End If
        End If
    Next c

    If testCol = 0 And Not ABSENT_IF_NO_TEST_COLUMN Then
        LogModelFailure label, Ar("r_notest")
        Exit Sub
    End If

    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        LogModelFailure label, Ar("r_nostudents")
        Exit Sub
    End If

    ' ---------- count students, detect conflicts ----------
    studentCount = CountStudents(wsSource, lastRow)
    UpdateCohortCount cohortName, studentCount, label

    cohortKey = SmartClean(cohortName)
    summaryKey = cohortKey & "|" & subjKey

    ' ---------- target sheet ----------
    s = ResolveTarget(cohortName)
    If s = 0 Then
        LogModelFailure label, Ar("r_nomap") & " " & cohortName
        Exit Sub
    End If
    If subjId = 0 Then
        LogModelFailure label, Ar("r_unknown") & ": " & subjectName
        Exit Sub
    End If

    ' ---------- duplicates ----------
    seenKey = CStr(s) & "|" & CStr(subjId)
    If g_Seen.Exists(seenKey) Then
        LogModelFailure label, Ar("r_dup") & " " & g_Seen(seenKey)
        Exit Sub
    End If

    If Not StoreSubject(wsSource, s, lastRow, testCol, subjId, subjectName, label) Then Exit Sub

    g_Seen.Add seenKey, label
    If fuzzyUsed Then
        g_Fuzzy = g_Fuzzy & Ar("n_fuzzy") & " " & subjectName & " " & ChrW(&H2190) & " " & SubjHeader(subjId) _
                  & "   [" & label & "]" & vbCrLf
    End If
    dictSummary(summaryKey) = studentCount
    If Not g_SubjectDisplay.Exists(subjKey) Then g_SubjectDisplay.Add subjKey, subjectName
End Sub

Private Function IsTestHeader(ByVal txt As String) As Boolean
    Dim t As String
    t = SmartClean(txt)
    IsTestHeader = (InStr(t, Ar("kw_test")) > 0 And InStr(t, "20") > 0)
End Function

Private Function CountStudents(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim dataArr As Variant
    Dim r As Long, cnt As Long
    dataArr = ws.Range(ws.Cells(DATA_START_ROW, 2), ws.Cells(lastRow, 2)).Value
    If IsArray(dataArr) Then
        For r = 1 To UBound(dataArr, 1)
            If SafeText(dataArr(r, 1)) <> "" Then cnt = cnt + 1
        Next r
    Else
        If SafeText(dataArr) <> "" Then cnt = 1
    End If
    CountStudents = cnt
End Function

Private Sub UpdateCohortCount(ByVal cohortName As String, ByVal studentCount As Long, ByVal label As String)
    Dim key As String
    key = SmartClean(cohortName)
    If Not g_CohortDisplay.Exists(key) Then g_CohortDisplay.Add key, cohortName
    If g_CohortCount.Exists(key) Then
        If g_CohortCount(key) <> studentCount Then
            g_CountConflicts = g_CountConflicts & _
                "  >> " & Ar("fl_cohort") & ": " & cohortName & vbCrLf & _
                "     " & Ar("fl_first") & ": " & g_CohortCount(key) & " (" & g_CohortOrigin(key) & ")" & vbCrLf & _
                "     " & Ar("fl_new") & ": " & studentCount & " (" & label & ")" & vbCrLf
        End If
    Else
        g_CohortCount.Add key, studentCount
        g_CohortOrigin.Add key, label
    End If
End Sub

' cohort name -> sheet index (0 = not mapped)
Private Function ResolveTarget(ByVal cohortName As String) As Long
    Dim c As String
    c = SmartClean(cohortName)
    If InStr(c, Ar("k_third")) = 0 Then Exit Function

    If InStr(c, Ar("k_adab")) > 0 Then ResolveTarget = 1: Exit Function
    If InStr(c, Ar("k_lughat")) > 0 And InStr(c, Ar("k_alman")) > 0 Then ResolveTarget = 2: Exit Function
    If InStr(c, Ar("k_ulum")) > 0 And InStr(c, Ar("k_tajrib")) > 0 Then
        Select Case LastDigit(c)
            Case "1": ResolveTarget = 3
            Case "2": ResolveTarget = 4
        End Select
        Exit Function
    End If
    If InStr(c, Ar("k_riyad")) > 0 Then ResolveTarget = 5: Exit Function
    If InStr(c, Ar("k_tasyir")) > 0 Then ResolveTarget = 6: Exit Function
End Function

Private Function LastDigit(ByVal s As String) As String
    Dim i As Long, code As Long
    For i = Len(s) To 1 Step -1
        code = AscW(Mid$(s, i, 1))
        If code >= 48 And code <= 57 Then
            LastDigit = Chr$(code)
            Exit Function
        End If
        If code >= 1632 And code <= 1641 Then
            LastDigit = Chr$(code - 1632 + 48)
            Exit Function
        End If
        If code >= 1776 And code <= 1785 Then
            LastDigit = Chr$(code - 1776 + 48)
            Exit Function
        End If
    Next i
End Function

' Stores one subject of one class in memory
Private Function StoreSubject(ByVal wsSource As Worksheet, ByVal s As Long, ByVal lastRow As Long, _
                              ByVal testCol As Long, ByVal subjId As Long, ByVal subjectName As String, _
                              ByVal label As String) As Boolean
    Dim dataArr As Variant, ids As Variant
    Dim outName() As String, outBirth() As Variant, outMark() As Variant
    Dim rowCount As Long, r As Long, n As Long, i As Long, j As Long, markCol As Long
    Dim nm As String, sig As String

    StoreSubject = False
    rowCount = lastRow - DATA_START_ROW + 1
    If rowCount <= 0 Then Exit Function

    markCol = testCol + MARK_OFFSET
    If testCol > 0 Then
        If markCol < 1 Or markCol > 8 Then
            LogModelFailure label, Ar("r_markcol")
            Exit Function
        End If
    End If

    dataArr = wsSource.Range(wsSource.Cells(DATA_START_ROW, 1), wsSource.Cells(lastRow, 8)).Value

    ReDim outName(1 To rowCount)
    ReDim outBirth(1 To rowCount)
    ReDim outMark(1 To rowCount)
    n = 0
    For r = 1 To UBound(dataArr, 1)
        nm = SafeText(dataArr(r, 2))
        If nm <> "" Then
            n = n + 1
            outName(n) = Trim$(nm & " " & SafeText(dataArr(r, 3)))
            If IsError(dataArr(r, 4)) Then outBirth(n) = "" Else outBirth(n) = dataArr(r, 4)
            If testCol > 0 Then
                outMark(n) = NormalizeMark(dataArr(r, markCol))
            Else
                outMark(n) = Ar("abs_text")
            End If
        End If
    Next r

    If n = 0 Then
        LogModelFailure label, Ar("r_nostudents")
        Exit Function
    End If
    If n > MAX_STUDENTS Then
        LogModelFailure label, n & " > " & MAX_STUDENTS & " : " & Ar("r_toomany")
        Exit Function
    End If

    ' ---------- student-list consistency within the class ----------
    sig = ""
    For i = 1 To n
        sig = sig & "|" & SmartClean(outName(i))
    Next i
    If mRoster(s) <> "" Then
        If mRoster(s) <> sig Then
            g_RosterConflicts = g_RosterConflicts & "  >> " & SheetName(s) & " - " & subjectName & " (" & label & ")" & vbCrLf
            If BLOCK_ON_ROSTER_MISMATCH Then Exit Function
        End If
    End If

    ' ---------- column of this subject inside the class ----------
    ids = SheetIds(s)
    j = 0
    For i = LBound(ids) To UBound(ids)
        If CLng(ids(i)) = subjId Then
            j = i - LBound(ids) + 1
            Exit For
        End If
    Next i
    If j = 0 Then
        LogModelFailure label, Ar("r_notinsheet") & " " & SheetName(s) & ": " & subjectName
        Exit Function
    End If

    If mN(s) = 0 Then
        mN(s) = n
        For i = 1 To n
            mNames(s, i) = outName(i)
            mBirth(s, i) = outBirth(i)
        Next i
        mRoster(s) = sig
    End If
    For i = 1 To mN(s)
        If i <= n Then mMarks(s, i, j) = outMark(i)
    Next i
    mLoaded(s, j) = True
    StoreSubject = True
End Function

Private Sub CheckCoverage()
    Dim s As Long, j As Long, ids As Variant, miss As String, txt As String
    For s = 1 To 6
        If mN(s) = 0 Then
            txt = txt & "  >> " & SheetName(s) & " : " & Ar("n_nodata") & vbCrLf
        Else
            ids = SheetIds(s)
            miss = ""
            For j = 1 To UBound(ids) + 1
                If Not mLoaded(s, j) Then
                    If miss <> "" Then miss = miss & " - "
                    miss = miss & SubjHeader(CLng(ids(j - 1)))
                End If
            Next j
            If miss <> "" Then txt = txt & "  >> " & SheetName(s) & " : " & miss & vbCrLf
        End If
    Next s
    If txt <> "" Then g_Notes = Ar("n_missing") & vbCrLf & txt
End Sub

Private Function NormalizeMark(ByVal v As Variant) As Variant
    Dim s As String, s2 As String
    If IsError(v) Then NormalizeMark = "": Exit Function
    If IsEmpty(v) Then NormalizeMark = "": Exit Function
    If VarType(v) <> vbString Then NormalizeMark = v: Exit Function
    s = Trim$(StripControl(CStr(v)))
    If s = "" Then NormalizeMark = "": Exit Function
    s2 = Replace(s, ",", ".")
    s2 = Replace(s2, ".", Application.International(xlDecimalSeparator))
    If IsNumeric(s2) Then
        NormalizeMark = CDbl(s2)
    Else
        NormalizeMark = s
    End If
End Function

'==============================================================================
'  Text normalization
'==============================================================================
Private Function StripControl(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, ChrW(8206), "")
    s = Replace(s, ChrW(8207), "")
    s = Replace(s, ChrW(1564), "")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    StripControl = s
End Function

Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeText = Trim$(StripControl(CStr(v)))
End Function

Private Function CleanLabel(ByVal s As String) As String
    Dim t As String, seps As String
    seps = " -" & ChrW(8211) & ChrW(8212) & ":|"
    t = Trim$(StripControl(s))
    Do While Len(t) > 0
        If InStr(seps, Left$(t, 1)) > 0 Then t = Mid$(t, 2) Else Exit Do
    Loop
    Do While Len(t) > 0
        If InStr(seps, Right$(t, 1)) > 0 Then t = Left$(t, Len(t) - 1) Else Exit Do
    Loop
    CleanLabel = t
End Function

Private Function SmartClean(ByVal txt As String) As String
    Dim t As String, k As Long, lamAlef As String
    lamAlef = ChrW(&H644) & ChrW(&H627)
    t = StripControl(txt)
    t = Replace(t, ChrW(1600), "")
    For k = 1611 To 1631
        t = Replace(t, ChrW(k), "")
    Next k
    For k = &HFEF5 To &HFEFC
        t = Replace(t, ChrW(k), lamAlef)
    Next k
    t = Trim$(t)
    t = Replace(t, " ", "")
    t = Replace(t, ".", "")
    t = Replace(t, ChrW(&H623), ChrW(&H627))
    t = Replace(t, ChrW(&H625), ChrW(&H627))
    t = Replace(t, ChrW(&H622), ChrW(&H627))
    t = Replace(t, ChrW(&H629), ChrW(&H647))
    t = Replace(t, ChrW(&H649), ChrW(&H64A))
    t = Replace(t, ChrW(&H6CC), ChrW(&H64A))
    t = Replace(t, ChrW(&H6A9), ChrW(&H643))
    SmartClean = t
End Function

' subject names: unifies physical-education variants (root / riyadiyya with heh / tarbiya+riyad)
' without swallowing mathematics (riyadiyyat keeps its own key)
Private Function SmartCleanSubject(ByVal txt As String) As String
    Dim t As String
    t = SmartClean(txt)
    If InStr(t, Ar("k_pe_root")) > 0 Or InStr(t, Ar("k_pe_ria")) > 0 _
       Or (InStr(t, Ar("k_pe2")) > 0 And InStr(t, Ar("k_pe3")) > 0) Then t = Ar("k_pe")
    SmartCleanSubject = t
End Function

'==============================================================================
'  PART 2 : build the output workbook
'==============================================================================
Private Sub BuildOutputWorkbook(ByVal dictSummary As Object)
    Dim ws As Worksheet, wsGen As Worksheet, wsSum As Worksheet
    Dim s As Long, nCols As Long

    Set g_wbOut = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsGen = g_wbOut.Worksheets(1)
    ResolveFonts wsGen

    For s = 1 To 6
        Application.StatusBar = Ar("working") & " " & SheetName(s)
        Set ws = g_wbOut.Worksheets.Add(After:=g_wbOut.Worksheets(g_wbOut.Worksheets.Count))
        BuildSeriesSheet ws, s
        nCols = LastDataCol(ws)
        ApplyLayout ws, nCols
        mLayoutNotes = mLayoutNotes & CheckLayout(ws, nCols)
    Next s

    Set wsSum = g_wbOut.Worksheets.Add(After:=g_wbOut.Worksheets(g_wbOut.Worksheets.Count))
    WriteSummarySheet wsSum, dictSummary
    BuildGen wsGen

    On Error Resume Next
    wsGen.Activate
    wsGen.Range("A1").Select
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
'  data of one class: computes averages, sorts, fills the table array
'------------------------------------------------------------------------------
Private Sub BuildSheetData(ByVal s As Long, ByRef data() As Variant, ByRef hdr() As Variant, _
                           ByRef hdrText() As String, ByRef nRows As Long, ByRef nCols As Long)
    Dim ids As Variant, coefs As Variant
    Dim m As Long, n As Long, i As Long, j As Long, p As Long, c As Long, t As Long
    Dim outv() As Variant, avgs() As Variant, key() As Double, idx() As Long
    Dim v As Variant, num As Double, den As Double, absTxt As String
    Dim pres As Long, pass As Long, exm As Long

    ids = SheetIds(s)
    coefs = SheetCoefs(s)
    m = UBound(ids) + 1
    nCols = 3 + m + 1
    n = mN(s)
    If n > 0 Then nRows = n Else nRows = 1

    ReDim hdr(1 To nCols)
    ReDim hdrText(1 To nCols)
    hdrText(1) = Ar("h_num")
    hdrText(2) = Ar("h_name")
    hdrText(3) = Ar("h_birth")
    For j = 1 To m
        hdrText(3 + j) = SubjHeader(CLng(ids(j - 1)))
    Next j
    hdrText(nCols) = Ar("h_avg")
    For c = 1 To nCols
        hdr(c) = hdrText(c)
    Next c

    ReDim outv(1 To nRows, 1 To m)
    ReDim avgs(1 To nRows)
    ReDim key(1 To nRows)
    ReDim idx(1 To nRows)
    ReDim data(1 To nRows, 1 To nCols)

    For i = 1 To n
        num = 0: den = 0: absTxt = ""
        For j = 1 To m
            If mLoaded(s, j) Then
                v = mMarks(s, i, j)
                If IsNum(v) Then
                    num = num + CDbl(coefs(j - 1)) * CDbl(v)
                    den = den + CDbl(coefs(j - 1))
                    outv(i, j) = CDbl(v)
                ElseIf VarType(v) = vbString Then
                    If Len(Trim$(CStr(v))) = 0 Then
                        ' blank cell: treated as exempt for this student only
                    ElseIf SmartClean(CStr(v)) = mExemptKey Then
                        outv(i, j) = Ar("exempt")
                        If CLng(ids(j - 1)) = 14 Then exm = exm + 1
                    Else
                        outv(i, j) = CStr(v)
                        If absTxt = "" Then absTxt = Trim$(CStr(v))
                    End If
                End If
            End If
        Next j
        If absTxt <> "" Then
            avgs(i) = absTxt
        ElseIf den > 0 Then
            avgs(i) = Application.WorksheetFunction.Round(num / den, 2)
        Else
            avgs(i) = Empty
        End If
        If IsNum(avgs(i)) Then
            key(i) = CDbl(avgs(i))
            If CDbl(avgs(i)) > 0 Then
                pres = pres + 1
                If CDbl(avgs(i)) >= PASS_MARK Then pass = pass + 1
            End If
        Else
            key(i) = -1
        End If
        idx(i) = i
    Next i

    ' stable descending sort by average
    For i = 2 To n
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If key(idx(j)) >= key(t) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next i

    For p = 1 To n
        i = idx(p)
        data(p, 1) = p
        data(p, 2) = mNames(s, i)
        v = mBirth(s, i)
        If VarType(v) = vbDate Then data(p, 3) = CDbl(v) Else data(p, 3) = v
        For j = 1 To m
            data(p, 3 + j) = outv(i, j)
        Next j
        data(p, nCols) = avgs(i)
    Next p

    mReg(s) = n
    mPres(s) = pres
    mPass(s) = pass
    mExempt(s) = exm
End Sub

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    IsNum = IsNumeric(v)
End Function

Private Function IsPresent(ByVal v As Variant) As Boolean
    If IsNum(v) Then
        If CDbl(v) > 0 Then IsPresent = True
    End If
End Function

'------------------------------------------------------------------------------
'  statistics of one column (present students only)
'------------------------------------------------------------------------------
Private Sub CalcColStat(ByRef vals() As Double, ByVal cnt As Long, ByRef st As ColStat)
    Dim i As Long, j As Long, k As Long, best As Long, bestCnt As Long
    Dim sum As Double, ss As Double, t As Double, m As Double, sd As Double
    Dim tmp() As Double, x As Double

    st.N = cnt
    If cnt = 0 Then Exit Sub

    st.MinV = vals(1): st.MaxV = vals(1)
    For i = 1 To cnt
        x = vals(i)
        sum = sum + x
        If x < st.MinV Then st.MinV = x
        If x > st.MaxV Then st.MaxV = x
        If x >= PASS_MARK Then st.Pass = st.Pass + 1
        If x < 7 Then
            st.Bands(0) = st.Bands(0) + 1
        ElseIf x < 9 Then
            st.Bands(1) = st.Bands(1) + 1
        ElseIf x < 10 Then
            st.Bands(2) = st.Bands(2) + 1
        ElseIf x < 12 Then
            st.Bands(3) = st.Bands(3) + 1
        ElseIf x < 14 Then
            st.Bands(4) = st.Bands(4) + 1
        ElseIf x < 16 Then
            st.Bands(5) = st.Bands(5) + 1
        ElseIf x < 18 Then
            st.Bands(6) = st.Bands(6) + 1
        Else
            st.Bands(7) = st.Bands(7) + 1
        End If
    Next i
    m = sum / cnt
    st.Mean = m

    For i = 1 To cnt
        ss = ss + (vals(i) - m) * (vals(i) - m)
    Next i
    st.Vari = ss / cnt
    sd = Sqr(st.Vari)
    st.Sd = sd

    ReDim tmp(1 To cnt)
    For i = 1 To cnt
        tmp(i) = vals(i)
    Next i
    For i = 2 To cnt
        t = tmp(i)
        j = i - 1
        Do While j >= 1
            If tmp(j) <= t Then Exit Do
            tmp(j + 1) = tmp(j)
            j = j - 1
        Loop
        tmp(j + 1) = t
    Next i
    If cnt Mod 2 = 1 Then
        st.Median = tmp((cnt + 1) \ 2)
    Else
        st.Median = (tmp(cnt \ 2) + tmp(cnt \ 2 + 1)) / 2
    End If

    best = 0: bestCnt = 1
    For i = 1 To cnt
        k = 0
        For j = 1 To cnt
            If Abs(vals(j) - vals(i)) < 0.000000001 Then k = k + 1
        Next j
        If k > bestCnt Then
            bestCnt = k
            best = i
        End If
    Next i
    If best > 0 Then st.ModeV = vals(best)

    For i = 1 To cnt
        x = vals(i)
        If x >= m + 1.5 * sd And x <= 20 Then st.Cls(1) = st.Cls(1) + 1
        If x >= m + 0.5 * sd And x < m + 1.5 * sd Then st.Cls(2) = st.Cls(2) + 1
        If x >= m - 0.5 * sd And x < m + 0.5 * sd Then st.Cls(3) = st.Cls(3) + 1
        If x >= m - 1.5 * sd And x < m - 0.5 * sd Then st.Cls(4) = st.Cls(4) + 1
        If x < m - 1.5 * sd Then st.Cls(5) = st.Cls(5) + 1
    Next i
End Sub

Private Sub ComputeStats(ByRef data() As Variant, ByVal nRows As Long, ByVal nCols As Long, ByRef st() As ColStat)
    Dim c As Long, i As Long, cnt As Long
    Dim vals() As Double
    ReDim st(FIRST_NUM_COL To nCols)
    ReDim vals(1 To nRows)
    For c = FIRST_NUM_COL To nCols
        cnt = 0
        For i = 1 To nRows
            If IsPresent(data(i, nCols)) Then
                If IsNum(data(i, c)) Then
                    cnt = cnt + 1
                    vals(cnt) = CDbl(data(i, c))
                End If
            End If
        Next i
        CalcColStat vals, cnt, st(c)
    Next c
End Sub

'------------------------------------------------------------------------------
'  one class sheet
'------------------------------------------------------------------------------
Private Sub BuildSeriesSheet(ByVal ws As Worksheet, ByVal s As Long)
    Dim data() As Variant, hdr() As Variant, hdrText() As String
    Dim st() As ColStat
    Dim nRows As Long, nCols As Long, lastRow As Long, c As Long, i As Long
    Dim area As Range, lo As ListObject
    Dim t1 As Long, tb As Long, hb As Long, hs As Long, lastUsed As Long
    Dim bands As Variant

    BuildSheetData s, data, hdr, hdrText, nRows, nCols
    ComputeStats data, nRows, nCols, st
    lastRow = FIRST_DATA_ROW + nRows - 1

    ws.Name = SheetName(s)
    ws.DisplayRightToLeft = True
    ws.Columns("A").ColumnWidth = ID_COL_WIDTH
    ws.Columns("C").ColumnWidth = DATE_COL_MIN_WIDTH
    ws.Range(ws.Columns(FIRST_NUM_COL), ws.Columns(nCols)).ColumnWidth = NUM_COL_WIDTH

    Set area = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, nCols))
    ws.Range(ws.Cells(FIRST_DATA_ROW, 2), ws.Cells(lastRow, 2)).NumberFormat = "@"
    ws.Range(ws.Cells(FIRST_DATA_ROW, 3), ws.Cells(lastRow, 3)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(FIRST_DATA_ROW, FIRST_NUM_COL), ws.Cells(lastRow, nCols)).NumberFormat = FMT_DEC

    ws.Cells(HEADER_ROW, 1).Resize(1, nCols).Value2 = hdr
    ws.Cells(FIRST_DATA_ROW, 1).Resize(nRows, nCols).Value2 = data

    With area
        .Font.Name = mFontBase: .Font.Size = 14: .Font.Bold = False
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Rows.RowHeight = DATA_ROW_HEIGHT
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    With ws.Range(ws.Cells(FIRST_DATA_ROW, 2), ws.Cells(lastRow, 2))
        .Font.Name = mFontName: .Font.Size = 14
        .ShrinkToFit = True
    End With
    With ws.Range(ws.Cells(FIRST_DATA_ROW, FIRST_NUM_COL), ws.Cells(lastRow, nCols))
        .Font.Name = mFontNum: .Font.Size = 15
    End With

    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=area, XlListObjectHasHeaders:=xlYes)
    lo.Name = "T_" & ws.Name
    lo.TableStyle = "": lo.ShowTotals = False
    lo.ShowTableStyleRowStripes = False: lo.ShowTableStyleColumnStripes = False: lo.ShowAutoFilter = True
    With lo.HeaderRowRange
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = HEADER_ROW_HEIGHT
    End With
    With ws.Range(ws.Cells(HEADER_ROW, FIRST_NUM_COL), ws.Cells(HEADER_ROW, nCols))
        .WrapText = True
        .Orientation = 90
    End With

    ' ---------- row 1 : title + counters ----------
    ws.Rows(TITLE_ROW).RowHeight = TITLE_ROW_HEIGHT
    MergeText ws, TITLE_ROW, 1, nCols - 6, Ar("title_main2") & " " & SeriesBranch(s)
    StyleBanner ws.Range(ws.Cells(TITLE_ROW, 1), ws.Cells(TITLE_ROW, nCols - 6))
    PutStripCell ws, TITLE_ROW, nCols - 5, nCols - 4, Ar("s_count"), nCols - 3, mReg(s)
    PutStripCell ws, TITLE_ROW, nCols - 2, nCols - 1, Ar("s_exempt"), nCols, mExempt(s)

    ' ---------- table 1 : averages / passes / rate / max / min ----------
    t1 = lastRow + GAP_BELOW_DATA + 1
    StyleGrid ws.Range(ws.Cells(t1 + 1, 1), ws.Cells(t1 + 5, nCols))
    PutHeaderRow ws, t1, nCols, "", hdrText
    PutLabel ws, t1 + 1, Ar("l_avg")
    PutLabel ws, t1 + 2, Ar("l_pass")
    PutLabel ws, t1 + 3, Ar("l_rate")
    PutLabel ws, t1 + 4, Ar("l_max")
    PutLabel ws, t1 + 5, Ar("l_min")
    For c = FIRST_NUM_COL To nCols
        If st(c).N > 0 Then
            ws.Cells(t1 + 1, c).Value2 = st(c).Mean
            ws.Cells(t1 + 2, c).Value2 = st(c).Pass
            ws.Cells(t1 + 3, c).Value2 = st(c).Pass / st(c).N
            ws.Cells(t1 + 4, c).Value2 = st(c).MaxV
            ws.Cells(t1 + 5, c).Value2 = st(c).MinV
        End If
    Next c
    ws.Range(ws.Cells(t1 + 1, FIRST_NUM_COL), ws.Cells(t1 + 1, nCols)).NumberFormat = FMT_DEC
    ws.Range(ws.Cells(t1 + 2, FIRST_NUM_COL), ws.Cells(t1 + 2, nCols)).NumberFormat = FMT_COUNT
    ws.Range(ws.Cells(t1 + 3, FIRST_NUM_COL), ws.Cells(t1 + 3, nCols)).NumberFormat = FMT_PCT
    ws.Range(ws.Cells(t1 + 4, FIRST_NUM_COL), ws.Cells(t1 + 5, nCols)).NumberFormat = FMT_DEC

    ' ---------- table 2 : score bands ----------
    tb = t1 + 5 + 2
    hb = tb + 1
    MergeText ws, tb, 1, nCols, Ar("t3_title")
    StyleTableTitle ws.Range(ws.Cells(tb, 1), ws.Cells(tb, nCols))
    ws.Rows(tb).RowHeight = TITLE_HEIGHT
    StyleGrid ws.Range(ws.Cells(hb + 1, 1), ws.Cells(hb + 8, nCols))
    PutHeaderRow ws, hb, nCols, Ar("t3_head"), hdrText
    bands = Array("07 >", "08.99 - 07", "09.99 - 09", "11.99 - 10", "13.99 - 12", "15.99 - 14", "17.99 - 16", "18 =<")
    For i = 0 To 7
        ws.Cells(hb + 1 + i, 1).NumberFormat = "@"
        PutLabel ws, hb + 1 + i, CStr(bands(i))
    Next i
    For c = FIRST_NUM_COL To nCols
        If st(c).N > 0 Then
            For i = 0 To 7
                ws.Cells(hb + 1 + i, c).Value2 = st(c).Bands(i)
            Next i
        End If
    Next c
    ws.Range(ws.Cells(hb + 1, FIRST_NUM_COL), ws.Cells(hb + 8, nCols)).NumberFormat = FMT_COUNT

    ' ---------- table 3 : statistical indicators + 5 standard classes ----------
    hs = hb + 8 + 2
    StyleGrid ws.Range(ws.Cells(hs + 1, 1), ws.Cells(hs + 8, nCols))
    StyleGrid ws.Range(ws.Cells(hs + 10, 1), ws.Cells(hs + 14, nCols))
    PutHeaderRow ws, hs, nCols, Ar("t2_head"), hdrText
    PutLabel ws, hs + 1, Ar("i_rate")
    PutLabel ws, hs + 2, Ar("i_mean")
    PutLabel ws, hs + 3, Ar("i_median")
    PutLabel ws, hs + 4, Ar("i_mode")
    PutLabel ws, hs + 5, Ar("i_range")
    PutLabel ws, hs + 6, Ar("i_sd")
    PutLabel ws, hs + 7, Ar("i_var")
    PutLabel ws, hs + 8, Ar("i_cv")
    ws.Rows(hs + 9).RowHeight = 5
    For i = 1 To 5
        PutLabel ws, hs + 9 + i, Ar("cls0" & CStr(i))
    Next i
    For c = FIRST_NUM_COL To nCols
        If st(c).N > 0 Then
            ws.Cells(hs + 1, c).Value2 = st(c).Pass / st(c).N
            ws.Cells(hs + 2, c).Value2 = st(c).Mean
            ws.Cells(hs + 3, c).Value2 = st(c).Median
            If Not IsEmpty(st(c).ModeV) Then ws.Cells(hs + 4, c).Value2 = st(c).ModeV
            ws.Cells(hs + 5, c).Value2 = st(c).MaxV - st(c).MinV
            ws.Cells(hs + 6, c).Value2 = st(c).Sd
            ws.Cells(hs + 7, c).Value2 = st(c).Vari
            If st(c).Mean <> 0 Then ws.Cells(hs + 8, c).Value2 = st(c).Sd / st(c).Mean
            For i = 1 To 5
                ws.Cells(hs + 9 + i, c).Value2 = st(c).Cls(i)
            Next i
        End If
    Next c
    ws.Range(ws.Cells(hs + 1, FIRST_NUM_COL), ws.Cells(hs + 1, nCols)).NumberFormat = FMT_PCT
    ws.Range(ws.Cells(hs + 2, FIRST_NUM_COL), ws.Cells(hs + 7, nCols)).NumberFormat = FMT_DEC
    ws.Range(ws.Cells(hs + 8, FIRST_NUM_COL), ws.Cells(hs + 8, nCols)).NumberFormat = FMT_PCT
    ws.Range(ws.Cells(hs + 10, FIRST_NUM_COL), ws.Cells(hs + 14, nCols)).NumberFormat = FMT_COUNT

    lastUsed = hs + 14
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastUsed, nCols)).Address
        .PrintTitleRows = "$1:$2"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
'  styles (copied from modBacEtab / modBacAnalysis)
'------------------------------------------------------------------------------
Private Sub StyleBanner(ByVal rng As Range)
    With rng
        .Font.Name = mFontBase: .Font.Size = 16: .Font.Bold = True
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
End Sub

Private Sub StyleTableTitle(ByVal rng As Range)
    With rng
        .Font.Name = mFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub StyleGrid(ByVal rng As Range)
    With rng
        .Font.Name = mFontNum: .Font.Size = 15: .Font.Bold = False
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        .Rows.RowHeight = DATA_ROW_HEIGHT
    End With
End Sub

Private Sub StyleHead(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = mFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub StyleLabel(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = mFontName: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub MergeText(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal txt As String)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        If c2 > c1 Then .Merge
        .Value2 = txt
    End With
End Sub

' header row of an analysis table: A:C merged label + subject headers as plain text
Private Sub PutHeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal avgCol As Long, _
                         ByVal lbl As String, ByRef hdrText() As String)
    Dim c As Long
    For c = FIRST_NUM_COL To avgCol
        ws.Cells(r, c).Value2 = hdrText(c)
    Next c
    StyleHead ws.Range(ws.Cells(r, 1), ws.Cells(r, avgCol))
    ws.Range(ws.Cells(r, FIRST_NUM_COL), ws.Cells(r, avgCol)).Orientation = 90
    MergeText ws, r, 1, 3, lbl
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Size = 16
    ws.Rows(r).RowHeight = HDR_ROW_HEIGHT
End Sub

Private Sub PutLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    MergeText ws, r, 1, 3, txt
    StyleLabel ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
End Sub

Private Sub PutStripCell(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, _
                         ByVal lbl As String, ByVal cv As Long, ByVal v As Long)
    MergeText ws, r, c1, c2, lbl
    StyleHead ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    With ws.Cells(r, cv)
        .Value2 = v
        .NumberFormat = FMT_COUNT
        .Font.Name = mFontNum: .Font.Size = 15: .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Function LastDataCol(ByVal ws As Worksheet) As Long
    LastDataCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
End Function

' widths/heights applied last so nothing earlier can override them
Private Sub ApplyLayout(ByVal ws As Worksheet, ByVal nCols As Long)
    If nCols < 4 Then Exit Sub
    ws.Columns(1).ColumnWidth = ID_COL_WIDTH
    ws.Columns(2).AutoFit
    ws.Columns(3).AutoFit
    If ws.Columns(3).ColumnWidth < DATE_COL_MIN_WIDTH Then ws.Columns(3).ColumnWidth = DATE_COL_MIN_WIDTH
    ws.Range(ws.Columns(FIRST_NUM_COL), ws.Columns(nCols)).ColumnWidth = NUM_COL_WIDTH
    ws.Rows(TITLE_ROW).RowHeight = TITLE_ROW_HEIGHT
    ws.Rows(HEADER_ROW).RowHeight = HEADER_ROW_HEIGHT
End Sub

Private Function CheckLayout(ByVal ws As Worksheet, ByVal nCols As Long) As String
    Dim c As Long, t As String
    If Abs(ws.Columns(1).ColumnWidth - ID_COL_WIDTH) > 0.2 Then
        t = t & vbLf & ws.Name & " : A = " & Format$(ws.Columns(1).ColumnWidth, "0.00")
    End If
    For c = FIRST_NUM_COL To nCols
        If Abs(ws.Columns(c).ColumnWidth - NUM_COL_WIDTH) > 0.2 Then
            t = t & vbLf & ws.Name & " : " & Replace(ws.Cells(1, c).Address(True, False), "$1", "") & _
                " = " & Format$(ws.Columns(c).ColumnWidth, "0.00")
        End If
    Next c
    If Abs(ws.Rows(TITLE_ROW).RowHeight - TITLE_ROW_HEIGHT) > 1 Then
        t = t & vbLf & ws.Name & " : row 1 = " & Format$(ws.Rows(TITLE_ROW).RowHeight, "0.0")
    End If
    If Abs(ws.Rows(HEADER_ROW).RowHeight - HEADER_ROW_HEIGHT) > 1 Then
        t = t & vbLf & ws.Name & " : row 2 = " & Format$(ws.Rows(HEADER_ROW).RowHeight, "0.0")
    End If
    CheckLayout = t
End Function

'==============================================================================
'  Gen sheet
'==============================================================================
Private Sub SectionTitle(ByVal g As Worksheet, ByVal r As Long, ByVal lastCol As Long, ByVal txt As String)
    MergeText g, r, 1, lastCol, txt
    With g.Range(g.Cells(r, 1), g.Cells(r, lastCol))
        .Font.Name = mFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    g.Rows(r).RowHeight = TITLE_HEIGHT
End Sub

Private Sub PutHeads(ByVal g As Worksheet, ByVal r As Long, ByVal heads As Variant)
    Dim i As Long
    For i = 0 To UBound(heads)
        g.Cells(r, 1 + i).Value2 = heads(i)
    Next i
    StyleHead g.Range(g.Cells(r, 1), g.Cells(r, 1 + UBound(heads)))
    g.Rows(r).RowHeight = GEN_HEAD_HEIGHT
End Sub

Private Sub StyleBlock(ByVal g As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal lastCol As Long)
    StyleGrid g.Range(g.Cells(r1, 1), g.Cells(r2, lastCol))
    With g.Range(g.Cells(r1, 1), g.Cells(r2, 1))
        .Font.Name = mFontName: .Font.Size = 14: .Font.Bold = True
        .ShrinkToFit = True
    End With
End Sub

Private Sub StyleTotalRow(ByVal g As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    With g.Range(g.Cells(r, 1), g.Cells(r, lastCol))
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
    End With
End Sub

Private Sub PutRate(ByVal g As Worksheet, ByVal r As Long, ByVal c As Long, ByVal pass As Long, ByVal pres As Long)
    If pres > 0 Then g.Cells(r, c).Value2 = pass / pres
End Sub

Private Sub BuildGen(ByVal g As Worksheet)
    Dim ord As Variant
    Dim tit(1 To 4) As Long, hd(1 To 4) As Long, fr(1 To 4) As Long, lr(1 To 4) As Long, tt(1 To 4) As Long
    Dim k As Long, s As Long, r As Long, j As Long, y As Long, yearStart As Long
    Dim tReg As Long, tPres As Long, tPass As Long

    ord = Array(1, 2, 5, 3, 4, 6)
    g.Name = "Gen"
    g.DisplayRightToLeft = True
    g.Cells.Font.Name = mFontBase
    g.Cells.Font.Size = 14
    g.Cells.VerticalAlignment = xlCenter
    g.Columns("A").ColumnWidth = 30
    g.Columns("B:J").ColumnWidth = 15

    For k = 0 To 5
        s = CLng(ord(k))
        tReg = tReg + mReg(s)
        tPres = tPres + mPres(s)
        tPass = tPass + mPass(s)
    Next k

    SectionTitle g, 1, 6, Ar("gen_title")

    ' ---------- 1 : registered / present / absent ----------
    tit(1) = 3: hd(1) = 4: fr(1) = 5: lr(1) = 10: tt(1) = 11
    SectionTitle g, tit(1), 4, Ar("g1")
    PutHeads g, hd(1), Array(Ar("h_branch"), Ar("h_reg"), Ar("h_pres"), Ar("h_abs"))
    StyleBlock g, fr(1), tt(1), 4
    For k = 0 To 5
        s = CLng(ord(k))
        r = fr(1) + k
        g.Cells(r, 1).Value2 = GenBranch(s)
        g.Cells(r, 2).Value2 = mReg(s)
        g.Cells(r, 3).Value2 = mPres(s)
        g.Cells(r, 4).Value2 = mReg(s) - mPres(s)
    Next k
    g.Cells(tt(1), 1).Value2 = Ar("total")
    g.Cells(tt(1), 2).Value2 = tReg
    g.Cells(tt(1), 3).Value2 = tPres
    g.Cells(tt(1), 4).Value2 = tReg - tPres
    StyleTotalRow g, tt(1), 4
    g.Range(g.Cells(fr(1), 2), g.Cells(tt(1), 4)).NumberFormat = FMT_COUNT

    ' ---------- 2 : general results ----------
    tit(2) = tt(1) + 2: hd(2) = tit(2) + 1: fr(2) = tit(2) + 2: lr(2) = fr(2) + 5: tt(2) = lr(2) + 1
    SectionTitle g, tit(2), 4, Ar("g2")
    PutHeads g, hd(2), Array(Ar("h_branch"), Ar("h_pres"), Ar("h_passed"), Ar("h_rate"))
    StyleBlock g, fr(2), tt(2), 4
    For k = 0 To 5
        s = CLng(ord(k))
        r = fr(2) + k
        g.Cells(r, 1).Value2 = GenBranch(s)
        g.Cells(r, 2).Value2 = mPres(s)
        g.Cells(r, 3).Value2 = mPass(s)
        PutRate g, r, 4, mPass(s), mPres(s)
    Next k
    g.Cells(tt(2), 1).Value2 = Ar("total")
    g.Cells(tt(2), 2).Value2 = tPres
    g.Cells(tt(2), 3).Value2 = tPass
    PutRate g, tt(2), 4, tPass, tPres
    StyleTotalRow g, tt(2), 4
    g.Range(g.Cells(fr(2), 2), g.Cells(tt(2), 3)).NumberFormat = FMT_COUNT
    g.Range(g.Cells(fr(2), 4), g.Cells(tt(2), 4)).NumberFormat = FMT_PCT

    ' ---------- 3 : first semester vs second semester ----------
    tit(3) = tt(2) + 2: hd(3) = tit(3) + 1: fr(3) = tit(3) + 2: lr(3) = fr(3) + 5: tt(3) = lr(3) + 1
    SectionTitle g, tit(3), 4, Ar("g3")
    PutHeads g, hd(3), Array(Ar("h_branch"), Ar("h_sem1"), Ar("h_sem2"), Ar("h_diff"))
    StyleBlock g, fr(3), tt(3), 4
    For k = 0 To 5
        s = CLng(ord(k))
        r = fr(3) + k
        g.Cells(r, 1).Value2 = GenBranch(s)
        PutRate g, r, 3, mPass(s), mPres(s)
    Next k
    g.Cells(tt(3), 1).Value2 = Ar("total")
    PutRate g, tt(3), 3, tPass, tPres
    StyleTotalRow g, tt(3), 4
    g.Range(g.Cells(fr(3), 2), g.Cells(tt(3), 4)).NumberFormat = FMT_PCT
    ' the only live formulas of the workbook: they depend on the manual column B
    For r = fr(3) To tt(3)
        g.Cells(r, 4).Formula = "=IF(OR(B" & r & "="""",C" & r & "=""""),"""",C" & r & "-B" & r & ")"
    Next r
    AddDiffFormat g, fr(3), tt(3), 4

    ' ---------- 4 : last five seasons ----------
    tit(4) = tt(3) + 2: hd(4) = tit(4) + 1: fr(4) = tit(4) + 2: lr(4) = fr(4) + 5: tt(4) = lr(4) + 1
    If Month(Date) >= SEASON_START_MONTH Then yearStart = Year(Date) Else yearStart = Year(Date) - 1
    SectionTitle g, tit(4), 6, Ar("g4")
    g.Cells(hd(4), 1).Value2 = Ar("h_branch")
    For j = 0 To 4
        y = yearStart - 4 + j
        g.Cells(hd(4), 2 + j).NumberFormat = "@"
        g.Cells(hd(4), 2 + j).Value2 = CStr(y) & "-" & CStr(y + 1)
    Next j
    StyleHead g.Range(g.Cells(hd(4), 1), g.Cells(hd(4), 6))
    g.Rows(hd(4)).RowHeight = GEN_HEAD_HEIGHT
    StyleBlock g, fr(4), tt(4), 6
    For k = 0 To 5
        s = CLng(ord(k))
        r = fr(4) + k
        g.Cells(r, 1).Value2 = GenBranch(s)
        PutRate g, r, 6, mPass(s), mPres(s)
    Next k
    g.Cells(tt(4), 1).Value2 = Ar("total")
    PutRate g, tt(4), 6, tPass, tPres
    StyleTotalRow g, tt(4), 6
    g.Range(g.Cells(fr(4), 2), g.Cells(tt(4), 6)).NumberFormat = FMT_PCT
    ' highest = green / lowest = red per row (manual past seasons); helpers in hidden H:I
    AddRowExtremesFormat g, fr(4), tt(4), 2, 6, 8, 9

    On Error Resume Next
    g.Activate
    ActiveWindow.DisplayGridlines = False
    With g.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Conditional formatting for Gen (same look as modCompaire: green = better,
' red = worse, white bold text).
' Locale-proof on purpose: every rule uses only operators and ABSOLUTE cell
' addresses (no function names, no list separators, no dependence on the
' active cell), so it works in any Excel language / regional setting.
' ---------------------------------------------------------------------------
Private Sub AddCF(ByVal rng As Range, ByVal formula As String, ByVal fillColor As Long, ByRef failed As Boolean)
    Dim fc As Object
    On Error Resume Next
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    If Err.Number <> 0 Or fc Is Nothing Then
        failed = True
        Err.Clear
    Else
        fc.Interior.Color = fillColor
        fc.Font.Color = RGB(255, 255, 255)
        fc.Font.Bold = True
        fc.StopIfTrue = False
        If Err.Number <> 0 Then
            failed = True
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub

' "difference" column of section 3: > 0 green, < 0 red, 0 or empty = no colour
Private Sub AddDiffFormat(ByVal g As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal col As Long)
    Dim r As Long, a As String, failed As Boolean
    g.Range(g.Cells(r1, col), g.Cells(r2, col)).FormatConditions.Delete
    For r = r1 To r2
        a = g.Cells(r, col).Address(True, True)
        AddCF g.Cells(r, col), "=(" & a & "<>"""")*(" & a & ">0)", RGB(39, 139, 82), failed
        AddCF g.Cells(r, col), "=(" & a & "<>"""")*(" & a & "<0)", RGB(190, 55, 55), failed
    Next r
    If failed Then mLayoutNotes = mLayoutNotes & vbLf & "Gen section 3: conditional formatting could not be applied."
End Sub

' Section 4: each row is compared on its own.
'   highest value  -> green, lowest value -> red
'   two cells tied at the highest (or lowest) -> both coloured
'   three or more tied, all values equal, or fewer than two values -> nothing
' Two helper cells per row (hMax, hMin) hold the value to highlight, or "n" when
' nothing must be highlighted; they are real formulas, so they follow manual input.
' The helper columns are hidden.
Private Sub AddRowExtremesFormat(ByVal g As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                 ByVal c1 As Long, ByVal c2 As Long, _
                                 ByVal hMax As Long, ByVal hMin As Long)
    Dim r As Long, c As Long, a As String, rowRng As String
    Dim mx As String, mn As String, ha As String, hb As String, failed As Boolean

    g.Range(g.Cells(r1, c1), g.Cells(r2, c2)).FormatConditions.Delete
    For r = r1 To r2
        rowRng = g.Range(g.Cells(r, c1), g.Cells(r, c2)).Address(True, True)
        mx = "MAX(" & rowRng & ")"
        mn = "MIN(" & rowRng & ")"
        g.Cells(r, hMax).Formula = "=IF(AND(COUNT(" & rowRng & ")>=2," & mx & "<>" & mn & _
                                   ",COUNTIF(" & rowRng & "," & mx & ")<=2)," & mx & ",""n"")"
        g.Cells(r, hMin).Formula = "=IF(AND(COUNT(" & rowRng & ")>=2," & mx & "<>" & mn & _
                                   ",COUNTIF(" & rowRng & "," & mn & ")<=2)," & mn & ",""n"")"
        ha = g.Cells(r, hMax).Address(True, True)
        hb = g.Cells(r, hMin).Address(True, True)
        For c = c1 To c2
            a = g.Cells(r, c).Address(True, True)
            AddCF g.Cells(r, c), "=(" & a & "<>"""")*(" & a & "=" & ha & ")", RGB(39, 139, 82), failed
            AddCF g.Cells(r, c), "=(" & a & "<>"""")*(" & a & "=" & hb & ")", RGB(190, 55, 55), failed
        Next c
    Next r
    g.Range(g.Columns(hMax), g.Columns(hMin)).Hidden = True
    If failed Then mLayoutNotes = mLayoutNotes & vbLf & "Gen section 4: conditional formatting could not be applied."
End Sub

'==============================================================================
'  Data-transfer summary sheet (kept for review) + processing notes
'==============================================================================
Private Sub WriteSummarySheet(ByVal ws As Worksheet, ByVal dictSummary As Object)
    Dim k As Variant, parts As Variant, outArr() As Variant
    Dim i As Long, n As Long, r As Long
    Dim keyCohort As String, keySubject As String
    Dim lines As Variant, ln As Variant, txt As String

    ws.Name = Ar("sum_name")
    ws.DisplayRightToLeft = True
    ws.Cells.Font.Name = mFontName
    ws.Cells.Font.Size = 14

    ws.Range("A1:C1").Value2 = Array(Ar("sum_c1"), Ar("sum_c2"), Ar("sum_c3"))
    StyleHead ws.Range("A1:C1")
    ws.Rows(1).RowHeight = GEN_HEAD_HEIGHT

    n = dictSummary.Count
    r = 1
    If n > 0 Then
        ReDim outArr(1 To n, 1 To 3)
        i = 0
        For Each k In dictSummary.Keys
            i = i + 1
            parts = Split(CStr(k), "|")
            keyCohort = parts(0)
            keySubject = parts(1)
            If g_CohortDisplay.Exists(keyCohort) Then outArr(i, 1) = g_CohortDisplay(keyCohort) Else outArr(i, 1) = keyCohort
            If g_SubjectDisplay.Exists(keySubject) Then outArr(i, 2) = g_SubjectDisplay(keySubject) Else outArr(i, 2) = keySubject
            outArr(i, 3) = dictSummary(k)
        Next k
        ws.Range("A2").Resize(n, 3).Value2 = outArr
        With ws.Range("A2:C" & (n + 1))
            .Font.Name = mFontName: .Font.Size = 14
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .ReadingOrder = xlRTL
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
            .Rows.RowHeight = DATA_ROW_HEIGHT
        End With
        ws.Range("A1:C" & (n + 1)).Sort _
            Key1:=ws.Range("A1"), Order1:=xlAscending, _
            Key2:=ws.Range("B1"), Order2:=xlAscending, _
            Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        ws.Range("A1:C" & (n + 1)).Columns.AutoFit
        r = n + 1
    End If

    ' ---------- processing notes (full text, nothing truncated) ----------
    txt = ""
    If g_OpenFailures <> "" Then txt = txt & Ar("fl_open") & vbCrLf & g_OpenFailures & vbCrLf
    If g_ModelFailures <> "" Then txt = txt & Ar("fl_model") & vbCrLf & g_ModelFailures & vbCrLf
    If g_RosterConflicts <> "" Then txt = txt & Ar("fl_roster") & vbCrLf & g_RosterConflicts & vbCrLf
    If g_CountConflicts <> "" Then txt = txt & Ar("fl_count") & vbCrLf & g_CountConflicts & vbCrLf
    If g_Fuzzy <> "" Then txt = txt & g_Fuzzy
    If g_Notes <> "" Then txt = txt & g_Notes & vbCrLf
    If txt <> "" Then
        r = r + 2
        ws.Cells(r, 1).Value2 = Ar("sum_log")
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 16
        lines = Split(txt, vbCrLf)
        For Each ln In lines
            If Len(Trim$(CStr(ln))) > 0 Then
                r = r + 1
                ws.Cells(r, 1).NumberFormat = "@"
                ws.Cells(r, 1).Value2 = CStr(ln)
                ws.Cells(r, 1).HorizontalAlignment = xlRight
                ws.Cells(r, 1).ReadingOrder = xlRTL
            End If
        Next ln
    End If
End Sub

'==============================================================================
'  Fonts (same logic as modBacEtab)
'==============================================================================
Private Sub ResolveFonts(ByVal ws As Worksheet)
    mFontNotes = ""
    mFontBase = PickFont(ws, FONT_BASE, FONT_SAFE_TEXT)
    mFontName = PickFont(ws, FONT_NAME, FONT_BASE, FONT_SAFE_TEXT)
    mFontNum = PickFont(ws, FONT_NUM, FONT_SAFE_NUM, FONT_SAFE_TEXT)
    ws.Cells.Clear
End Sub

Private Function PickFont(ByVal ws As Worksheet, ParamArray names() As Variant) As String
    Dim i As Long, first As String, chosen As String
    first = CStr(names(LBound(names)))
    If Not CHECK_FONTS Then
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
        If InStr(mFontNotes, first & " ->") = 0 Then mFontNotes = mFontNotes & vbLf & first & " -> " & chosen
    End If
End Function

Private Function FontAvailable(ByVal ws As Worksheet, ByVal fontName As String) As Boolean
    Dim samples(1 To 2) As String, i As Long, wBase As Double, wTest As Double
    samples(1) = "MWmwIl0123456789 The quick brown fox " & ArHex("627 644 639 631 628 64A 629 20 645 62F 631 633 629")
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
'  Files
'==============================================================================
Private Function FolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(p) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then FolderExists = False
    Err.Clear
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir$(fullPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function

Private Function PathInUse(ByVal fullPath As String) As Boolean
    Dim book As Workbook
    If FileExists(fullPath) Then
        PathInUse = True
        Exit Function
    End If
    For Each book In Application.Workbooks
        If StrComp(book.FullName, fullPath, vbTextCompare) = 0 Then
            PathInUse = True
            Exit Function
        End If
    Next book
End Function

' first file: <stem>.xlsx   then <stem>_01.xlsx, <stem>_02.xlsx ...
Private Function NewOutputPath(ByVal folder As String, ByVal stem As String) As String
    Dim fullPath As String, n As Long
    fullPath = folder & Application.PathSeparator & stem & ".xlsx"
    n = 0
    Do While PathInUse(fullPath)
        n = n + 1
        fullPath = folder & Application.PathSeparator & stem & "_" & Format$(n, "00") & ".xlsx"
    Loop
    NewOutputPath = fullPath
End Function

'==============================================================================
'  Output folder (created next to the tool workbook when missing)
'==============================================================================
Private Function OutputFolderName() As String
    OutputFolderName = ArHex("0645 0639 0627 0644 062C 0629 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0641 0635 0644 0020 0030 0032")
End Function

Private Function EnsureFolder(ByVal p As String) As Boolean
    If Right$(p, 1) = Application.PathSeparator Then p = Left$(p, Len(p) - 1)
    If Not FolderExists(p) Then
        On Error Resume Next
        MkDir p
        Err.Clear
        On Error GoTo 0
    End If
    EnsureFolder = FolderExists(p)
End Function

'==============================================================================
'  Arabic text (hex code points)
'==============================================================================
Private Function ArHex(ByVal codes As String) As String
    Dim token As Variant
    For Each token In Split(codes, " ")
        If Len(CStr(token)) > 0 Then ArHex = ArHex & ChrW(CLng("&H" & CStr(token)))
    Next token
End Function

Private Function Ar(ByVal key As String) As String
    Select Case key
        Case "working": Ar = ArHex("62C 627 631 64D 20 627 644 62A 62D 636 64A 631 3A")
        Case "fonts": Ar = ArHex("62A 645 20 627 633 62A 628 62F 627 644 20 62E 637 648 637 20 63A 64A 631 20 645 62A 648 641 631 629 20 628 62E 637 648 637 20 628 62F 64A 644 629 3A")
        Case "layout_warn": Ar = ArHex("62A 646 628 64A 647 3A 20 62A 639 630 631 20 636 628 637 20 627 644 623 628 639 627 62F 20 627 644 62A 627 644 64A 629 20 643 645 627 20 647 648 20 645 637 644 648 628 3A")
        Case "sheets": Ar = ArHex("639 62F 62F 20 627 644 623 648 631 627 642")
        Case "h_num": Ar = ArHex("627 644 631 642 645")
        Case "h_name": Ar = ArHex("627 644 644 642 628 20 648 627 644 627 633 645")
        Case "h_birth": Ar = ArHex("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F")
        Case "h_avg": Ar = ArHex("627 644 645 639 62F 644")
        Case "done": Ar = ArHex("62A 645 20 625 646 634 627 621 20 645 644 641 20 627 644 62A 62D 644 64A 644 3A")
        Case "failed": Ar = ArHex("62A 639 630 631 20 625 646 634 627 621 20 645 644 641 20 627 644 62A 62D 644 64A 644 3A")
        Case "s_count": Ar = ArHex("62A 639 62F 627 62F 20 627 644 642 633 645")
        Case "l_avg": Ar = ArHex("645 639 62F 644 627 62A 20 627 644 645 648 627 62F")
        Case "l_pass": Ar = ArHex("645 62C 645 648 639 20 627 644 646 627 62C 62D 64A 646")
        Case "l_rate": Ar = ArHex("646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 643 644 20 645 627 62F 629")
        Case "l_max": Ar = ArHex("623 639 644 649 20 645 639 62F 644")
        Case "l_min": Ar = ArHex("623 62F 646 649 20 645 639 62F 644")
        Case "t2_head": Ar = ArHex("627 644 645 624 634 631 20 627 644 625 62D 635 627 626 64A")
        Case "i_rate": Ar = ArHex("646 633 628 629 20 627 644 646 62C 627 62D")
        Case "i_mean": Ar = ArHex("627 644 645 62A 648 633 637 20 627 644 62D 633 627 628 64A")
        Case "i_median": Ar = ArHex("627 644 648 633 64A 637")
        Case "i_mode": Ar = ArHex("627 644 645 646 648 627 644")
        Case "i_range": Ar = ArHex("627 644 645 62F 649")
        Case "i_sd": Ar = ArHex("627 644 627 646 62D 631 627 641 20 627 644 645 639 64A 627 631 64A")
        Case "i_var": Ar = ArHex("627 644 62A 628 627 64A 646")
        Case "i_cv": Ar = ArHex("645 639 627 645 644 20 627 644 627 62E 62A 644 627 641")
        Case "t3_title": Ar = ArHex("62A 648 632 64A 639 20 627 644 645 639 62F 644 627 62A 20 62D 633 628 20 627 644 641 626 627 62A")
        Case "t3_head": Ar = ArHex("627 644 641 626 627 62A 20 2F 20 627 644 645 648 627 62F")
        Case "gen_title": Ar = ArHex("625 62D 635 627 626 64A 627 62A 20 639 627 645 629")
        Case "g1": Ar = ArHex("31 2D 20 639 62F 62F 20 627 644 645 633 62C 644 64A 646")
        Case "g2": Ar = ArHex("32 2D 20 646 62A 627 626 62C 20 639 627 645 629")
        Case "h_branch": Ar = ArHex("627 644 634 639 628 629")
        Case "h_reg": Ar = ArHex("627 644 645 633 62C 644 64A 646")
        Case "h_pres": Ar = ArHex("627 644 62D 627 636 631 64A 646")
        Case "h_abs": Ar = ArHex("627 644 63A 627 626 628 64A 646")
        Case "h_passed": Ar = ArHex("627 644 646 627 62C 62D 64A 646")
        Case "h_rate": Ar = ArHex("646 633 628 629 20 627 644 646 62C 627 62D")
        Case "total": Ar = ArHex("627 644 645 62C 645 648 639 20 627 644 639 627 645")
        Case "sj1": Ar = ArHex("627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627")
        Case "sj2": Ar = ArHex("627 644 641 644 633 641 629")
        Case "sj3": Ar = ArHex("627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627")
        Case "sj4": Ar = ArHex("627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629")
        Case "sj5": Ar = ArHex("627 644 644 63A 629 20 627 644 641 631 646 633 64A 629")
        Case "sj6": Ar = ArHex("627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629")
        Case "sj7": Ar = ArHex("627 644 644 63A 629 20 627 644 623 62C 646 628 64A 629 20 627 644 62B 627 644 62B 629")
        Case "sj8": Ar = ArHex("627 644 631 64A 627 636 64A 627 62A")
        Case "sj9": Ar = ArHex("627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629")
        Case "sj10": Ar = ArHex("639 644 648 645 20 637 628 64A 639 64A 629 20 648 62D 64A 627 629")
        Case "sj11": Ar = ArHex("62A 2E 20 627 644 645 62D 627 633 628 64A 20 648 627 644 645 627 644 64A")
        Case "sj12": Ar = ArHex("627 644 627 642 62A 635 627 62F 20 648 627 644 645 646 627 62C 645 646 62A")
        Case "sj13": Ar = ArHex("627 644 642 627 646 648 646")
        Case "sj14": Ar = ArHex("62A 20 628 62F 646 64A 629 20 648 631 64A 627 636 64A 629")
        Case "app_title2": Ar = ArHex("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 627 62E 62A 628 627 631 627 62A")
        Case "out_name": Ar = ArHex("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 627 62E 62A 628 627 631 627 62A 20 30 32")
        Case "save_tool": Ar = ArHex("627 62D 641 638 20 645 644 641 20 627 644 623 62F 627 629 20 623 648 644 627 64B 20 641 64A 20 645 62C 644 62F 60C 20 62B 645 20 636 639 20 645 62C 644 62F 20 627 644 645 635 62F 631 20 628 62C 627 646 628 647 2E")
        Case "no_folder": Ar = ArHex("644 645 20 623 62C 62F 20 645 62C 644 62F 20 627 644 645 635 62F 631 20 628 62C 627 646 628 20 645 644 641 20 627 644 623 62F 627 629 3A")
        Case "no_files": Ar = ArHex("644 627 20 62A 648 62C 62F 20 645 644 641 627 62A 20 78 6C 73 78 20 641 64A 20 645 62C 644 62F 20 627 644 645 635 62F 631 3A")
        Case "see_sheet": Ar = ArHex("627 644 62A 641 627 635 64A 644 20 627 644 643 627 645 644 629 20 641 64A 20 648 631 642 629 20 AB 646 642 644 20 627 644 628 64A 627 646 627 62A BB 2E")
        Case "ok_all": Ar = ArHex("639 648 644 62C 62A 20 643 644 20 627 644 645 644 641 627 62A 20 628 646 62C 627 62D 20 62F 648 646 20 62A 639 627 631 636 627 62A 2E")
        Case "fl_already": Ar = ArHex("28 645 641 62A 648 62D 20 641 64A 20 45 78 63 65 6C 20 2D 20 623 63A 644 642 647 20 648 623 639 62F 20 627 644 645 62D 627 648 644 629 29")
        Case "kw_cohort": Ar = ArHex("627 644 641 648 62C 20 627 644 62A 631 628 648 64A")
        Case "kw_subject": Ar = ArHex("645 627 62F 629")
        Case "kw_al": Ar = ArHex("627 644")
        Case "kw_test": Ar = ArHex("627 644 627 62E 62A 628 627 631")
        Case "r_a5": Ar = ArHex("644 645 20 623 62C 62F 20 627 644 641 648 62C 20 623 648 20 627 644 645 627 62F 629 20 641 64A 20 627 644 62E 644 64A 629 20 41 35")
        Case "r_notest": Ar = ArHex("644 645 20 623 62C 62F 20 639 645 648 62F 20 627 644 627 62E 62A 628 627 631 20 2F 32 30 20 641 64A 20 627 644 639 645 648 62F 64A 646 20 47 20 648 20 48")
        Case "r_nostudents": Ar = ArHex("644 627 20 62A 648 62C 62F 20 635 641 648 641 20 62A 644 627 645 64A 630")
        Case "r_nomap": Ar = ArHex("627 644 641 648 62C 20 63A 64A 631 20 645 631 62A 628 637 20 628 623 64A 20 648 631 642 629 3A")
        Case "r_unknown": Ar = ArHex("645 627 62F 629 20 63A 64A 631 20 645 639 631 648 641 629")
        Case "r_dup": Ar = ArHex("645 643 631 631 20 644 646 641 633 20 627 644 641 648 62C 20 648 627 644 645 627 62F 629 20 2D 20 633 628 642 20 62A 62D 645 64A 644 647 20 645 646")
        Case "r_markcol": Ar = ArHex("639 645 648 62F 20 627 644 646 642 627 637 20 62E 627 631 62C 20 627 644 646 637 627 642")
        Case "r_toomany": Ar = ArHex("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 64A 62A 62C 627 648 632 20 627 644 62D 62F 20 627 644 645 633 645 648 62D")
        Case "r_notinsheet": Ar = ArHex("627 644 645 627 62F 629 20 644 64A 633 62A 20 636 645 646 20 645 648 627 62F 20 627 644 648 631 642 629")
        Case "fl_cohort": Ar = ArHex("627 644 641 648 62C")
        Case "fl_first": Ar = ArHex("623 648 644 20 62A 633 62C 64A 644")
        Case "fl_new": Ar = ArHex("627 644 639 62F 62F 20 627 644 62C 62F 64A 62F")
        Case "k_third": Ar = ArHex("62B 627 644 62B 647")
        Case "k_adab": Ar = ArHex("627 62F 627 628")
        Case "k_lughat": Ar = ArHex("644 63A 627 62A")
        Case "k_alman": Ar = ArHex("627 644 645 627 646 64A 647")
        Case "k_ulum": Ar = ArHex("639 644 648 645")
        Case "k_tajrib": Ar = ArHex("62A 62C 631 64A 628 64A 647")
        Case "k_riyad": Ar = ArHex("631 64A 627 636 64A 627 62A")
        Case "k_tasyir": Ar = ArHex("62A 633 64A 64A 631")
        Case "k_amaz": Ar = ArHex("627 645 627 632 64A 63A")
        Case "k_basic": Ar = ArHex("627 633 627 633 64A 647")
        Case "k_pe": Ar = ArHex("628 62F 646 64A 647")
        Case "k_pe_root": Ar = ArHex("628 62F 646")
        Case "k_pe_ria": Ar = ArHex("631 64A 627 636 64A 647")
        Case "n_fuzzy": Ar = ArHex("62A 645 20 627 644 62A 639 631 641 20 639 644 649 20 627 644 645 627 62F 629 20 631 63A 645 20 62E 637 623 20 625 645 644 627 626 64A 3A")
        Case "k_pe2": Ar = ArHex("62A 631 628 64A 647")
        Case "k_pe3": Ar = ArHex("631 64A 627 636 64A")
        Case "k_l3": Ar = ArHex("62B 627 644 62B")
        Case "k_ger": Ar = ArHex("627 644 645 627 646 64A")
        Case "k_spa": Ar = ArHex("627 633 628 627 646 64A")
        Case "k_ita": Ar = ArHex("627 64A 637 627 644 64A")
        Case "k_ar": Ar = ArHex("639 631 628 64A")
        Case "k_philo": Ar = ArHex("641 644 633 641")
        Case "k_hist": Ar = ArHex("62A 627 631 64A 62E")
        Case "k_geo": Ar = ArHex("62C 63A 631 627 641")
        Case "k_isl": Ar = ArHex("627 633 644 627 645")
        Case "k_fr": Ar = ArHex("641 631 646 633")
        Case "k_en1": Ar = ArHex("627 646 62C 644 64A 632")
        Case "k_en2": Ar = ArHex("627 646 643 644 64A 632")
        Case "k_phys": Ar = ArHex("641 64A 632 64A")
        Case "k_svt": Ar = ArHex("637 628 64A 639")
        Case "k_acc": Ar = ArHex("645 62D 627 633 628")
        Case "k_eco": Ar = ArHex("627 642 62A 635 627 62F")
        Case "k_law": Ar = ArHex("642 627 646 648 646")
        Case "k_math": Ar = ArHex("631 64A 627 636 64A 627 62A")
        Case "n_nodata": Ar = ArHex("644 645 20 64A 64F 62D 645 64E 651 644 20 623 64A 20 645 635 62F 631 20 644 647 630 627 20 627 644 642 633 645")
        Case "n_missing": Ar = ArHex("62A 646 628 64A 647 20 2D 20 645 648 627 62F 20 628 644 627 20 645 635 62F 631 20 28 639 64F 62F 651 62A 20 645 639 641 649 64B 20 645 646 647 627 20 627 644 642 633 645 20 643 627 645 644 627 64B 20 648 644 627 20 64A 62F 62E 644 20 645 639 627 645 644 647 627 20 641 64A 20 627 644 645 639 62F 644 29 3A")
        Case "exempt": Ar = ArHex("645 639 641 649")
        Case "abs_text": Ar = ArHex("63A 20 645")
        Case "bn1": Ar = ArHex("622 62F 627 628 20 648 641 644 633 641 629")
        Case "bn2": Ar = ArHex("644 63A 627 62A 20 623 62C 646 628 64A 629")
        Case "bn3": Ar = ArHex("639 644 648 645 20 62A 62C 631 64A 628 64A 629 20 30 31")
        Case "bn4": Ar = ArHex("639 644 648 645 20 62A 62C 631 64A 628 64A 629 20 30 32")
        Case "bn5": Ar = ArHex("631 64A 627 636 64A 627 62A")
        Case "bn6": Ar = ArHex("62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F")
        Case "gb1": Ar = ArHex("33 20 622 62F 627 628 20 648 641 644 633 641 629")
        Case "gb2": Ar = ArHex("33 20 644 63A 627 62A 20 623 62C 646 628 64A 629")
        Case "gb3": Ar = ArHex("33 20 639 644 648 645 20 62A 62C 20 30 31")
        Case "gb4": Ar = ArHex("33 20 639 644 648 645 20 62A 62C 20 30 32")
        Case "gb5": Ar = ArHex("33 20 631 64A 627 636 64A 627 62A")
        Case "gb6": Ar = ArHex("33 20 62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F")
        Case "title_main2": Ar = ArHex("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 62E 62A 628 627 631 627 62A 20 627 644 641 635 644 20 627 644 62B 627 646 64A 20 2D 20 642 633 645 20 62B 627 644 62B 629")
        Case "s_exempt": Ar = ArHex("627 644 645 639 641 64A 648 646")
        Case "cls01": Ar = ArHex("627 644 641 626 629 20 30 31")
        Case "cls02": Ar = ArHex("627 644 641 626 629 20 30 32")
        Case "cls03": Ar = ArHex("627 644 641 626 629 20 30 33")
        Case "cls04": Ar = ArHex("627 644 641 626 629 20 30 34")
        Case "cls05": Ar = ArHex("627 644 641 626 629 20 30 35")
        Case "g3": Ar = ArHex("33 2D 20 645 642 627 631 646 629 20 646 633 628 629 20 627 644 646 62C 627 62D 20 628 64A 646 20 627 644 641 635 644 64A 646 20 627 644 623 648 644 20 648 627 644 62B 627 646 64A")
        Case "g4": Ar = ArHex("34 2D 20 645 642 627 631 646 629 20 646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 627 644 641 635 644 20 627 644 62B 627 646 64A 20 644 644 645 648 627 633 645 20 627 644 645 627 636 64A 629")
        Case "h_sem1": Ar = ArHex("627 644 641 635 644 20 627 644 623 648 644")
        Case "h_sem2": Ar = ArHex("627 644 641 635 644 20 627 644 62B 627 646 64A")
        Case "h_diff": Ar = ArHex("627 644 641 627 631 642")
        Case "sum_name": Ar = ArHex("646 642 644 20 627 644 628 64A 627 646 627 62A")
        Case "sum_c1": Ar = ArHex("627 644 641 648 62C 20 627 644 62A 631 628 648 64A")
        Case "sum_c2": Ar = ArHex("627 644 645 627 62F 629")
        Case "sum_c3": Ar = ArHex("639 62F 62F 20 627 644 637 644 627 628")
        Case "sum_log": Ar = ArHex("645 644 627 62D 638 627 62A 20 627 644 645 639 627 644 62C 629")
        Case "fl_open": Ar = ArHex("645 644 641 627 62A 20 62A 639 630 631 20 641 62A 62D 647 627 3A")
        Case "fl_model": Ar = ArHex("645 644 641 627 62A 2F 623 648 631 627 642 20 644 645 20 62A 64F 646 642 644 3A")
        Case "fl_roster": Ar = ArHex("642 627 626 645 629 20 627 644 62A 644 627 645 64A 630 20 62A 62E 62A 644 641 20 639 646 20 627 644 645 62D 645 64E 651 644 629 20 633 627 628 642 627 64B 20 644 644 641 648 62C 20 646 641 633 647 20 28 644 645 20 62A 64F 646 642 644 20 627 644 645 627 62F 629 29 3A")
        Case "fl_count": Ar = ArHex("62A 646 628 64A 647 3A 20 627 62E 62A 644 627 641 20 641 64A 20 639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 644 644 641 648 62C 20 646 641 633 647 3A")
    End Select
End Function
