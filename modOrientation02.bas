Attribute VB_Name = "modOrientation02"
Option Explicit

Private m_runSucceeded As Boolean

' ============================================================
' Orientation 02: science (Data) and arts (Data01) in Term02Guidance.xlsx
' Original21/15-column tables and two wishes per track; builds one live general report.
' Run order: Run_Transfer02 -> Run_StreamPreferences02 -> Run_Orientation02
' Science section column N is ignored; arts has no section column.
' ============================================================

Private Const DST_FILE As String = "Term02Guidance.xlsx"
Private Const DST_SHEET As String = "Data"
Private Const DST_TABLE As String = "Data"
Private Const FIRST_ROW As Long = 3

' table columns
Private Const C_BIRTH As Long = 3     ' C
Private Const C_REPEAT As Long = 5    ' E
Private Const C_MA As Long = 7        ' G math
Private Const C_PH As Long = 8        ' H physics
Private Const C_NS As Long = 9        ' I natural sciences
Private Const C_AVG As Long = 13      ' M term average
Private Const C_W1 As Long = 15       ' O wish 1
Private Const C_W2 As Long = 16       ' P wish 2
Private Const C_P1 As Long = 17       ' Q..T profiles (math, tech, exp, mgmt)
Private Const C_OUT As Long = 21      ' U suggestion
Private Const NC As Long = 21

' ---- parameters (edit here) ----
' Flexible ceilings: branches may remain below their share; no minimum is forced.
' Mathematics and technical mathematics are assigned only when requested.
Private Const PASS_MARK As Double = 10
Private Const TOP_PCT As Double = 0.05        ' automatic wish 1; nearest integer, halves up
Private Const CAP_MATH As Double = 0.18       ' seat count rounded down from share of passers
Private Const CAP_TECH As Double = 0.22       ' seat count rounded down
Private Const CAP_EXP As Double = 0.55        ' experimental sciences: round up
Private Const CAP_MGMT As Double = 1          ' management: round up, absorbs all remaining passers
Private Const AGE_OFFSET As Long = 17         ' born in (working year - 17) or later => repeat the year
Private Const WORK_YEAR_OVERRIDE As Long = 0  ' 0 = use Year(Date)
Private Const COMPAT_GAP As Double = 1.0      ' max advantage of the actual alternative over wish 1
Private Const COMPAT_MIN As Double = 12.99    ' all-profile exemption from subject conditions (experimental sciences and management only)
Private Const COMPAT_FOR_EXP As Boolean = True
Private Const EPS As Double = 0.000001

Private Const D_ORIENT As Long = 5
Private Const D_REPEAT As Long = 6

' ---- in-memory data ----
Private m_n As Long
Private m_prof() As Double
Private m_hasP() As Boolean
Private m_sj() As Double
Private m_hasS() As Boolean
Private m_avg() As Double
Private m_w1() As Long
Private m_w2() As Long
Private g_br(1 To 4) As String
Private g_f(1 To 4) As String
Private g_yes As String

' ---- statistics for the report ----
Private m_sWish(1 To 4, 1 To 2) As Long
Private m_sDec(1 To 4) As Long
Private m_sMet(1 To 4) As Long
Private m_nStud As Long, m_nPass As Long, m_nRep As Long, m_nOri As Long

' One public entry point runs both tracks and saves only after both succeed.
Public Sub Run_Orientation02()
    Dim wb As Workbook, opened As Boolean, ws As Worksheet, lo As ListObject
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldEvents As Boolean
    Dim failed As Boolean, stateSaved As Boolean, errText As String, msg As String
    Dim scienceTouched As Boolean, artsTouched As Boolean
    Dim v As Variant, scienceSheet As Worksheet, artsSheet As Worksheet
    m_runSucceeded = False
    On Error GoTo EH
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    stateSaved = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    InitText
    Set wb = OpenWb(DST_FILE, False, opened)
    If wb.ReadOnly Then Err.Raise vbObjectError + 2, , DST_FILE & " is open read-only."

    ' Missing destinations are skipped; present, wholly empty tables are valid.
    Set scienceSheet = FindWorksheet(wb, "Data")
    Set artsSheet = FindWorksheet(wb, "Data01")
    If scienceSheet Is Nothing And artsSheet Is Nothing Then Err.Raise vbObjectError + 84, , "Neither 'Data' nor 'Data01' exists in " & DST_FILE & ". Run data transfer first."
    If Not scienceSheet Is Nothing Then
        Set lo = GuidanceTable(scienceSheet, 21)
        scienceSheet.Calculate
        If Not GuidanceTableIsEmpty(lo) Then
            v = lo.DataBodyRange.Value2
            ValidateGuidanceData v, wb.Date1904, lo.DataBodyRange.Row
        End If
    End If
    If Not artsSheet Is Nothing Then
        Set lo = GuidanceTable(artsSheet, 15)
        artsSheet.Calculate
        If Not GuidanceTableIsEmpty(lo) Then
            v = lo.DataBodyRange.Value2
            ValidateArtsData v, wb.Date1904, lo.DataBodyRange.Row
        End If
    End If
    If Not scienceSheet Is Nothing Then
        scienceTouched = True
        CheckTrackWrite wb, scienceSheet
    End If
    If Not artsSheet Is Nothing Then
        artsTouched = True
        CheckTrackWrite wb, artsSheet
    End If
    If Not scienceSheet Is Nothing Then msg = OrientScience(wb) Else msg = "Data: missing destination; skipped."
    If Not artsSheet Is Nothing Then msg = msg & vbLf & vbLf & OrientArts(wb) Else msg = msg & vbLf & "Data01: missing destination; skipped."
    BuildGeneralReport wb
    Application.Calculation = xlCalculationAutomatic
    wb.Save
    m_runSucceeded = True
Done:
    On Error Resume Next
    If failed And opened Then
        wb.Close SaveChanges:=False
    ElseIf failed And Not wb Is Nothing Then
        For Each ws In wb.Worksheets
            If (ws.Name = "Data" And scienceTouched) Or (ws.Name = "Data01" And artsTouched) Then
                Set lo = Nothing
                If ws.Name = "Data01" Then Set lo = GuidanceTable(ws, 15) Else Set lo = GuidanceTable(ws, 21)
                If Not lo Is Nothing Then LockAndProtect ws, lo.DataBodyRange
            End If
        Next ws
    End If
    If stateSaved Then
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
    End If
    On Error GoTo 0
    If failed Then MsgBox "Orientation failed:" & vbLf & errText, vbCritical, "Orientation 02" Else MsgBox msg, vbInformation, "Orientation 02"
    Exit Sub
EH:
    m_runSucceeded = False
    failed = True
    errText = Err.Description
    If (scienceTouched Or artsTouched) And Not opened Then
        errText = errText & vbLf & vbLf & ArtsText("failure_unsaved")
    End If
    Resume Done
End Sub

Private Function ArtsWish(ByVal x As Variant) As Long
    Dim t As String
    If IsError(x) Then Exit Function
    t = Norm(x)
    If Len(t) = 0 Then Exit Function
    If InStr(t, Norm(ArtsText("f_arts"))) > 0 Or InStr(t, Norm(ArtsText("f_philo"))) > 0 Then
        ArtsWish = 1
    ElseIf InStr(t, Norm(ArtsText("f_lang"))) > 0 Then
        ArtsWish = 2
    End If
End Function

Private Sub ValidateArtsData(ByRef v As Variant, ByVal date1904 As Boolean, ByVal firstRow As Long)
    Dim i As Long, c As Long, w1 As Long, w2 As Long, d As Date, location As String
    Dim problems As Object
    Set problems = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(v, 1)
        location = "Data01!" & (firstRow + i - 1)
        If IsError(v(i, 2)) Or BlankValue(v(i, 2)) Then AddValidationIssue problems, location, "B"
        For c = 6 To 10
            If Not IsNum(v(i, c)) Then AddValidationIssue problems, location, ColumnLetter(c)
        Next c
        For c = 13 To 14
            If Not IsNum(v(i, c)) Then AddValidationIssue problems, location, ColumnLetter(c)
        Next c
        If Not BlankValue(v(i, 3)) Then
            If Not TryBirthDate(v(i, 3), date1904, d) Then AddValidationIssue problems, location, "C"
        End If
        w1 = ArtsWish(v(i, 11)): w2 = ArtsWish(v(i, 12))
        If IsError(v(i, 11)) Or (Not BlankValue(v(i, 11)) And w1 = 0) Then AddValidationIssue problems, location, "K"
        If IsError(v(i, 12)) Or (Not BlankValue(v(i, 12)) And w2 = 0) Then AddValidationIssue problems, location, "L"
        If w1 > 0 And w1 = w2 Then AddValidationIssue problems, location, "K=L"
    Next i
    RaiseValidationIssues problems, "Letters=invalid/missing columns; K=L=repeated non-empty wishes."
End Sub

' Language candidates: wish 1, profile conversions, then unfulfilled wish 2.
' A fulfilled arts first wish is not displaced unless the language advantage > 1.00.
Private Function ArtsTier(ByVal firstWish As Long, ByVal secondWish As Long, ByVal languageProfile As Double, ByVal artsProfile As Double) As Long
    If firstWish = 2 Then
        ArtsTier = 1
    ElseIf firstWish = 1 Or (firstWish = 0 And secondWish <> 2) Then
        If languageProfile - artsProfile > COMPAT_GAP + EPS Then ArtsTier = 2
    ElseIf firstWish = 0 And secondWish = 2 Then
        ArtsTier = 3
    End If
End Function

Private Function ArtsAhead(ByVal i As Long, ByVal j As Long, ByRef v As Variant) As Boolean
    Dim a As Double, b As Double
    a = CDbl(v(i, 13)): b = CDbl(v(j, 13))
    If a <> b Then
        ArtsAhead = a > b
        Exit Function
    End If
    a = (CDbl(v(i, 8)) + CDbl(v(i, 9))) / 2
    b = (CDbl(v(j, 8)) + CDbl(v(j, 9))) / 2
    If a <> b Then
        ArtsAhead = a > b
        Exit Function
    End If
    ArtsAhead = CDbl(v(i, 10)) > CDbl(v(j, 10))
End Function

Private Sub SortArtsCandidates(ByRef candidates() As Long, ByVal count As Long, ByRef v As Variant)
    Dim i As Long, j As Long, t As Long
    For i = 2 To count
        t = candidates(i): j = i - 1
        Do While j >= 1
            If ArtsAhead(t, candidates(j), v) Then
                candidates(j + 1) = candidates(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        candidates(j + 1) = t
    Next i
End Sub

Private Function OrientArts(ByVal wb As Workbook) As String
    Const TOP_ARTS As Double = 0.05
    Const NORMAL_LANG As Double = 0.2
    Const MAX_LANG As Double = 0.25
    Dim ws As Worksheet, lo As ListObject, v As Variant, output() As Variant
    Dim decision() As Long, w1() As Long, w2() As Long, elite() As Boolean
    Dim passers() As Long, candidates() As Long
    Dim n As Long, nPass As Long, nRep As Long, nOri As Long, nIssue As Long
    Dim i As Long, j As Long, b As Long, tier As Long, nCand As Long, topK As Long
    Dim countLang As Long, eliteLang As Long, hardCap As Long, normalCap As Long, tierCap As Long
    Dim workYear As Long, cutoff As Long, bYear As Long
    Dim wishes(1 To 2) As Long, counts(1 To 2) As Long, met(1 To 2) As Long
    Set ws = wb.Worksheets("Data01")
    Set lo = GuidanceTable(ws, 15)
    If GuidanceTableIsEmpty(lo) Then
        LockAndProtect ws, Nothing
        OrientArts = "Arts: no registered pupils."
        Exit Function
    End If
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(10).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With
    v = lo.DataBodyRange.Value2
    n = UBound(v, 1)
    ReDim output(1 To n, 1 To 1)
    ReDim decision(1 To n)
    ReDim w1(1 To n): ReDim w2(1 To n): ReDim elite(1 To n)
    ReDim passers(1 To n): ReDim candidates(1 To n)
    If WORK_YEAR_OVERRIDE > 0 Then workYear = WORK_YEAR_OVERRIDE Else workYear = Year(Date)
    cutoff = workYear - AGE_OFFSET
    For i = 1 To n
        w1(i) = ArtsWish(v(i, 11)): w2(i) = ArtsWish(v(i, 12))
        If w1(i) > 0 Then wishes(w1(i)) = wishes(w1(i)) + 1
        If CDbl(v(i, 10)) >= PASS_MARK - EPS Then
            nPass = nPass + 1
            passers(nPass) = i
            decision(i) = 1
        ElseIf IsRepeater(v(i, 5)) Then
            decision(i) = D_ORIENT
        Else
            bYear = BirthYear(v(i, 3), wb.Date1904)
            If bYear = 0 Then
                nIssue = nIssue + 1
            ElseIf bYear >= cutoff Then
                decision(i) = D_REPEAT
            Else
                decision(i) = D_ORIENT
            End If
        End If
    Next i

    ' Excel sorted the full table by the term average before creating passers().
    topK = ShareCountUp(TOP_ARTS, nPass)
    hardCap = ShareCount(MAX_LANG, nPass)
    normalCap = ShareCountUp(NORMAL_LANG, nPass)
    If normalCap > hardCap Then normalCap = hardCap
    For j = 1 To topK
        i = passers(j)
        If w1(i) > 0 Then
            elite(i) = True
            decision(i) = w1(i)
            If decision(i) = 2 Then countLang = countLang + 1
        End If
    Next j
    eliteLang = countLang
    ' For cohorts of 1..3 passers, rounding can make top5% exceed a zero-seat cap.
    ' The explicit unconditional elite entitlement takes precedence; no other seat is added.
    If hardCap < eliteLang Then hardCap = eliteLang

    For tier = 1 To 3
        ' Conversions stop at the normal TOTAL language count; requests may use flexibility.
        tierCap = hardCap
        If tier = 2 Then tierCap = normalCap
        nCand = 0
        For j = 1 To nPass
            i = passers(j)
            If Not elite(i) Then
                ' Inclusive 10.00 threshold: BOTH French and English are required.
                If CDbl(v(i, 8)) >= PASS_MARK And CDbl(v(i, 9)) >= PASS_MARK Then
                    If ArtsTier(w1(i), w2(i), CDbl(v(i, 13)), CDbl(v(i, 14))) = tier Then
                        nCand = nCand + 1
                        candidates(nCand) = i
                    End If
                End If
            End If
        Next j
        If nCand > 1 Then SortArtsCandidates candidates, nCand, v
        For j = 1 To nCand
            If countLang >= tierCap Then Exit For
            decision(candidates(j)) = 2
            countLang = countLang + 1
        Next j
    Next tier

    For i = 1 To n
        Select Case decision(i)
        Case 1, 2
            b = decision(i)
            counts(b) = counts(b) + 1
            If w1(i) = b Then met(b) = met(b) + 1
            If b = 1 Then output(i, 1) = ArtsText("arts") Else output(i, 1) = ArtsText("lang")
        Case D_ORIENT
            nOri = nOri + 1
            output(i, 1) = L("orient")
        Case D_REPEAT
            nRep = nRep + 1
            output(i, 1) = L("repeat")
        End Select
    Next i
    If counts(1) + counts(2) <> nPass Then Err.Raise vbObjectError + 82, , "Arts orientation totals do not match."
    If nPass + nRep + nOri + nIssue <> n Then Err.Raise vbObjectError + 82, , "Arts classification totals do not match."
    If counts(2) > hardCap Then Err.Raise vbObjectError + 82, , "Language capacity exceeded."
    lo.DataBodyRange.Columns(15).Value2 = output
    LockAndProtect ws, lo.DataBodyRange
    OrientArts = "Arts: " & n & " students, " & nPass & " passers." & vbLf & _
                 ArtsText("arts") & ": " & counts(1) & "; " & ArtsText("lang") & ": " & counts(2) & _
                 " (20% seats: " & normalCap & "; flexible max: " & hardCap & ")."
    If nIssue > 0 Then OrientArts = OrientArts & vbLf & nIssue & " row(s) left blank: missing birth date."
End Function

Private Function OrientScience(ByVal wbDst As Workbook) As String
    Dim wsDst As Worksheet, lo As ListObject, body As Range
    Dim v As Variant, out() As Variant, x As Variant
    Dim dec() As Long, met() As Long, pidx() As Long, cand() As Long
    Dim isPass() As Boolean, issue() As Boolean
    Dim cnt(1 To 4) As Long, cap(1 To 4) As Long
    Dim i As Long, j As Long, k As Long, b As Long, r As Long, alternative As Long
    Dim nPass As Long, nIssue As Long, nCand As Long, topK As Long
    Dim workYear As Long, cutoff As Long, bYear As Long
    Dim subjCols As Variant
    Dim msg As String

    ResetStats
    Set wsDst = wbDst.Worksheets("Data")
    Set lo = GuidanceTable(wsDst, 21)
    If lo.ListColumns.Count < NC Then
        Err.Raise vbObjectError + 6, , "Table '" & DST_TABLE & "' must have at least " & NC & " columns (A:U)."
    End If
    If GuidanceTableIsEmpty(lo) Then
        ResetStats
        LockAndProtect wsDst, Nothing
        OrientScience = "Science: no registered pupils."
        Exit Function
    End If
    If lo.Range.Column <> 1 Or lo.HeaderRowRange.Row <> 2 Then
        Err.Raise vbObjectError + 6, , "Table '" & DST_TABLE & "' must start at A2."
    End If
    wsDst.Calculate
    v = lo.DataBodyRange.Value2

    CheckTrackWrite wbDst, wsDst

    ' show every row, then sort by term average (descending)
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(C_AVG).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

    Set body = lo.DataBodyRange
    v = body.Value2
    m_n = UBound(v, 1)

    ReDim m_prof(1 To m_n, 1 To 4)
    ReDim m_hasP(1 To m_n, 1 To 4)
    ReDim m_sj(1 To m_n, 1 To 3)
    ReDim m_hasS(1 To m_n, 1 To 3)
    ReDim m_avg(1 To m_n)
    ReDim m_w1(1 To m_n)
    ReDim m_w2(1 To m_n)
    ReDim dec(1 To m_n)
    ReDim met(1 To m_n)
    ReDim isPass(1 To m_n)
    ReDim issue(1 To m_n)
    ReDim out(1 To m_n, 1 To 1)

    If WORK_YEAR_OVERRIDE > 0 Then workYear = WORK_YEAR_OVERRIDE Else workYear = Year(Date)
    cutoff = workYear - AGE_OFFSET
    subjCols = Array(C_MA, C_PH, C_NS)

    ' ---------- classify ----------
    For i = 1 To m_n
        m_w1(i) = WishBranch(v(i, C_W1))
        m_w2(i) = WishBranch(v(i, C_W2))
        For k = 1 To 4
            x = v(i, C_P1 + k - 1)
            If IsNum(x) Then
                m_prof(i, k) = CDbl(x)
                m_hasP(i, k) = True
            End If
        Next k
        For k = 1 To 3
            x = v(i, subjCols(k - 1))
            If IsNum(x) Then
                m_sj(i, k) = CDbl(x)
                m_hasS(i, k) = True
            End If
        Next k

        If IsNum(v(i, C_AVG)) Then
            m_avg(i) = CDbl(v(i, C_AVG))
            If m_avg(i) >= PASS_MARK - EPS Then
                isPass(i) = True
                nPass = nPass + 1
            ElseIf IsRepeater(v(i, C_REPEAT)) Then
                dec(i) = D_ORIENT
            Else
                bYear = BirthYear(v(i, C_BIRTH), wbDst.Date1904)
                If bYear = 0 Then
                    issue(i) = True
                    nIssue = nIssue + 1
                ElseIf bYear >= cutoff Then
                    dec(i) = D_REPEAT
                Else
                    dec(i) = D_ORIENT
                End If
            End If
        Else
            issue(i) = True
            nIssue = nIssue + 1
        End If
    Next i

    ' ---------- passers ----------
    If nPass > 0 Then
        ReDim pidx(1 To nPass)
        ReDim cand(1 To nPass)
        j = 0
        For i = 1 To m_n
            If isPass(i) Then
                j = j + 1
                pidx(j) = i
            End If
        Next i
        SortPass pidx, nPass

        cap(1) = ShareCount(CAP_MATH, nPass)
        cap(2) = ShareCount(CAP_TECH, nPass)
        cap(3) = ShareCountUp(CAP_EXP, nPass)
        cap(4) = ShareCountUp(CAP_MGMT, nPass)

        ' Top 5%: grant a non-empty first wish automatically, before all rounds.
        ' No subject, profile-compatibility or capacity test is applied here.
        ' Their seats still count toward the branch ceiling in later rounds.
        topK = ShareCountNearest(TOP_PCT, nPass)
        For j = 1 To topK
            i = pidx(j)
            b = m_w1(i)
            If b > 0 Then
                dec(i) = b
                met(i) = 1
                cnt(b) = cnt(b) + 1
            End If
        Next j

        ' rounds: wish 1 for everybody, then wish 2
        For r = 1 To 2
            For b = 1 To 4
                nCand = 0
                For j = 1 To nPass
                    i = pidx(j)
                    If dec(i) = 0 Then
                        If WishOf(i, r) = b Then
                            If Eligible(i, b, r) Then
                                nCand = nCand + 1
                                cand(nCand) = i
                            End If
                        End If
                    End If
                Next j
                If nCand > 1 Then SortCand cand, nCand, b
                For j = 1 To nCand
                    If cnt(b) >= cap(b) Then Exit For
                    i = cand(j)
                    alternative = 0
                    If r = 1 Then alternative = ProposedAlternative(i, cnt, cap)
                    If Eligible(i, b, r, alternative) Then
                        dec(i) = b
                        met(i) = r
                        cnt(b) = cnt(b) + 1
                    End If
                Next j
            Next b
        Next r

        ' Leftovers: only experimental sciences or management, with subject checks.
        For j = 1 To nPass
            i = pidx(j)
            If dec(i) = 0 Then
                b = BestRemaining(i, cnt, cap)
                If b = 0 Then
                    Err.Raise vbObjectError + 56, , "No eligible branch at row " & (body.Row + i - 1) & ". No suggestion was written."
                End If
                dec(i) = b
                cnt(b) = cnt(b) + 1
            End If
        Next j

        ' Seat changes can make the actual alternative differ from the earlier one.
        ' Reconsider wish 1 against the branch actually assigned, while seats remain.
        RestoreFirstWishes pidx, nPass, cand, dec, met, cnt, cap
    End If

    ' ---------- statistics + output ----------
    m_nStud = m_n
    m_nPass = nPass
    For i = 1 To m_n
        If m_w1(i) > 0 Then m_sWish(m_w1(i), 1) = m_sWish(m_w1(i), 1) + 1
        If m_w2(i) > 0 Then m_sWish(m_w2(i), 2) = m_sWish(m_w2(i), 2) + 1
        Select Case dec(i)
        Case 1 To 4
            m_sDec(dec(i)) = m_sDec(dec(i)) + 1
            If m_w1(i) = dec(i) Then m_sMet(dec(i)) = m_sMet(dec(i)) + 1
        Case D_ORIENT
            m_nOri = m_nOri + 1
        Case D_REPEAT
            m_nRep = m_nRep + 1
        End Select
        out(i, 1) = DecText(dec(i))
    Next i

    If m_sDec(1) + m_sDec(2) + m_sDec(3) + m_sDec(4) <> nPass Then
        Err.Raise vbObjectError + 57, , "Orientation count does not match the number of passers."
    End If
    If m_nPass + m_nOri + m_nRep + nIssue <> m_nStud Then
        Err.Raise vbObjectError + 57, , "Student classification totals do not match."
    End If
    For b = 1 To 4
        If cnt(b) > cap(b) Then Err.Raise vbObjectError + 57, , "Branch capacity was exceeded."
    Next b

    body.Columns(C_OUT).Value2 = out
    LockAndProtect wsDst, body

    msg = "Done. Students: " & m_n & ", passers: " & nPass & vbLf & _
          g_br(1) & ": " & m_sDec(1) & " / max " & cap(1) & vbLf & _
          g_br(2) & ": " & m_sDec(2) & " / max " & cap(2) & vbLf & _
          g_br(3) & ": " & m_sDec(3) & " / max " & cap(3) & vbLf & _
          g_br(4) & ": " & m_sDec(4) & vbLf & _
          L("repeat") & ": " & m_nRep & "   " & L("orient") & ": " & m_nOri
    If nIssue > 0 Then msg = msg & vbLf & vbLf & nIssue & " row(s) left blank (average or birth date unreadable)."

    OrientScience = msg
End Function

' ---------- rules ----------
Private Function WishOf(ByVal i As Long, ByVal r As Long) As Long
    If r = 1 Then WishOf = m_w1(i) Else WishOf = m_w2(i)
End Function

Private Function NeedCompat(ByVal b As Long) As Boolean
    Select Case b
    Case 1, 2, 4: NeedCompat = True
    Case 3: NeedCompat = COMPAT_FOR_EXP
    Case Else: NeedCompat = False
    End Select
End Function

Private Function Eligible(ByVal i As Long, ByVal b As Long, ByVal r As Long, Optional ByVal alternative As Long = 0) As Boolean
    If b < 1 Or b > 4 Then Exit Function
    If Not m_hasP(i, b) Then Exit Function
    ' The explicit three-subject rule for requested experimental sciences remains.
    If b = 3 Then
        If RequestedExperimental(i, r) Then
            ' Exception: when wish 2 is mathematics or technical mathematics and was proposed as the
            ' real alternative (subject conditions and capacity already checked in ProposedAlternative),
            ' wish 1 is refused if that profile exceeds the experimental profile by more than COMPAT_GAP.
            ' The pupil is then directed by wish 2 in round 2.
            If r = 1 And (alternative = 1 Or alternative = 2) Then
                If Not Compat(i, 3, alternative) Then Exit Function
            End If
            Eligible = True
            Exit Function
        End If
    End If
    ' Retain the all-high-profile exemption from first-wish subject conditions.
    ' Its gap is measured only against the actual proposed alternative.
    If r = 1 Then
        If FlexibleFirstWish(i, b, alternative) Then
            Eligible = True
            Exit Function
        End If
    End If
    If Not Allowed(i, b) Then Exit Function
    If r = 1 And NeedCompat(b) Then
        If Not Compat(i, b, alternative) Then Exit Function
    End If
    Eligible = True
End Function

' A request in round 1 or 2 qualifies with math, physics and NS all >= 10.
' Existing G:H:I marks are used; this module does not calculate annual averages.
' Round 2 is reached only for pupils whose first wish was not assigned.
Private Function RequestedExperimental(ByVal i As Long, ByVal r As Long) As Boolean
    Dim k As Long
    If r < 1 Or r > 2 Then Exit Function
    If WishOf(i, r) <> 3 Then Exit Function
    For k = 1 To 3
        If Not m_hasS(i, k) Then Exit Function
        If m_sj(i, k) < PASS_MARK Then Exit Function
    Next k
    RequestedExperimental = True
End Function

' All-high-profile subject exemption, using the actual alternative.
' It applies ONLY to experimental sciences and management: the subject conditions of
' mathematics and technical mathematics can never be bypassed.
' Zero alternative is a candidate precheck; allocation supplies a real proposal.
Private Function FlexibleFirstWish(ByVal i As Long, ByVal b As Long, ByVal alternative As Long) As Boolean
    Dim k As Long
    If b < 3 Or b > 4 Then Exit Function
    If m_w1(i) <> b Then Exit Function
    For k = 1 To 4
        If Not m_hasP(i, k) Then Exit Function
        If m_prof(i, k) <= COMPAT_MIN + EPS Then Exit Function
    Next k
    FlexibleFirstWish = Compat(i, b, alternative)
End Function

' Subject conditions per branch (1 math, 2 tech, 3 experimental, 4 management)
Private Function Allowed(ByVal i As Long, ByVal b As Long) As Boolean
    Dim ma As Double, ph As Double, ns As Double, lowest As Double
    Select Case b
    Case 4
        Allowed = m_hasP(i, 4)
    Case 1
        If m_hasS(i, 1) And m_hasS(i, 2) Then
            Allowed = (m_sj(i, 1) >= 10 And m_sj(i, 2) >= 10)
        End If
    Case 2
        If m_hasS(i, 1) And m_hasS(i, 2) Then
            ma = m_sj(i, 1): ph = m_sj(i, 2)
            If ma < ph Then lowest = ma Else lowest = ph
            Allowed = Not ((ma < 10 And ph < 10) Or (lowest < 9 And Rnd2(ma + ph) < 20))
        End If
    Case 3
        If m_hasS(i, 1) And m_hasS(i, 2) And m_hasS(i, 3) Then
            ma = m_sj(i, 1): ph = m_sj(i, 2): ns = m_sj(i, 3)
            Allowed = Not ((ns < 10 And ma < 10) Or (ns < 10 And ph < 10) Or (ma < 10 And ph < 10 And ns < 11))
        End If
    End Select
End Function

' Compare wish 1 with the actual proposed branch only, never the highest profile.
' An alternative must exceed wish 1 by more than 1.00 to reject it on profiles.
Private Function Compat(ByVal i As Long, ByVal b As Long, ByVal alternative As Long) As Boolean
    If b < 1 Or b > 4 Then Exit Function
    If Not m_hasP(i, b) Then Exit Function
    If alternative = 0 Then
        ' No feasible alternative, or a candidate precheck before allocation.
        Compat = True
        Exit Function
    End If
    If alternative < 1 Or alternative > 4 Then Exit Function
    If Not m_hasP(i, alternative) Then Exit Function
    Compat = (m_prof(i, alternative) - m_prof(i, b) <= COMPAT_GAP + EPS)
End Function

' The real alternative is wish 2 when eligible and available; otherwise fallback.
' BestRemaining considers only experimental sciences and management with seats.
Private Function ProposedAlternative(ByVal i As Long, ByRef counts() As Long, ByRef capacities() As Long) As Long
    Dim b As Long
    b = m_w2(i)
    If b >= 1 And b <= 4 Then
        If counts(b) < capacities(b) Then
            If Eligible(i, b, 2) Then
                ProposedAlternative = b
                Exit Function
            End If
        End If
    End If
    ProposedAlternative = BestRemaining(i, counts, capacities)
End Function

' Reconcile against actual assigned branches after both rounds and fallback.
' Moves only toward a pupil's first wish; a pupil can therefore move at most once.
' Repeated passes allow a released seat to satisfy another first wish.
Private Sub RestoreFirstWishes(ByRef passers() As Long, ByVal nPass As Long, ByRef candidates() As Long, _
                               ByRef decisions() As Long, ByRef met() As Long, ByRef counts() As Long, ByRef capacities() As Long)
    Dim b As Long, i As Long, j As Long, nCand As Long, previous As Long
    Dim changed As Boolean
    Do
        changed = False
        For b = 1 To 4
            nCand = 0
            If counts(b) < capacities(b) Then
                For j = 1 To nPass
                    i = passers(j)
                    If m_w1(i) = b And decisions(i) <> b Then
                        If decisions(i) >= 1 And decisions(i) <= 4 Then
                            If Eligible(i, b, 1, decisions(i)) Then
                                nCand = nCand + 1
                                candidates(nCand) = i
                            End If
                        End If
                    End If
                Next j
                If nCand > 1 Then SortCand candidates, nCand, b
                For j = 1 To nCand
                    If counts(b) >= capacities(b) Then Exit For
                    i = candidates(j)
                    previous = decisions(i)
                    counts(previous) = counts(previous) - 1
                    decisions(i) = b
                    met(i) = 1
                    counts(b) = counts(b) + 1
                    changed = True
                Next j
            End If
        Next b
    Loop While changed
End Sub

' Ranking inside a branch: profile, then math, physics, natural sciences, then term average
Private Function Ahead(ByVal i As Long, ByVal j As Long, ByVal b As Long) As Boolean
    Dim a As Double, c As Double, k As Long
    a = -1000000#: c = -1000000#
    If m_hasP(i, b) Then a = m_prof(i, b)
    If m_hasP(j, b) Then c = m_prof(j, b)
    If a <> c Then
        Ahead = (a > c)
        Exit Function
    End If
    For k = 1 To 3
        a = -1000000#: c = -1000000#
        If m_hasS(i, k) Then a = m_sj(i, k)
        If m_hasS(j, k) Then c = m_sj(j, k)
        If a <> c Then
            Ahead = (a > c)
            Exit Function
        End If
    Next k
    If m_avg(i) <> m_avg(j) Then Ahead = (m_avg(i) > m_avg(j))
End Function

Private Sub SortCand(ByRef c() As Long, ByVal cnt As Long, ByVal b As Long)
    Dim i As Long, j As Long, t As Long
    For i = 2 To cnt
        t = c(i)
        j = i - 1
        Do While j >= 1
            If Ahead(t, c(j), b) Then
                c(j + 1) = c(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        c(j + 1) = t
    Next i
End Sub

Private Sub SortPass(ByRef p() As Long, ByVal cnt As Long)
    Dim i As Long, j As Long, t As Long
    For i = 2 To cnt
        t = p(i)
        j = i - 1
        Do While j >= 1
            If m_avg(p(j)) < m_avg(t) Then
                p(j + 1) = p(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        p(j + 1) = t
    Next i
End Sub

' ---------- parsing ----------
Private Sub InitText()
    g_br(1) = L("math"): g_br(2) = L("tech"): g_br(3) = L("exp"): g_br(4) = L("mgmt")
    g_f(1) = Norm(L("f_math")): g_f(2) = Norm(L("f_tech"))
    g_f(3) = Norm(L("f_exp")): g_f(4) = Norm(L("f_mgmt"))
    g_yes = Norm(L("yes"))
End Sub

' wish text -> branch number (0 = blank / unrecognized)
Private Function WishBranch(ByVal x As Variant) As Long
    Dim t As String
    If IsError(x) Then Exit Function
    t = Norm(x)
    If Len(t) = 0 Then Exit Function
    If InStr(t, g_f(2)) > 0 Then
        WishBranch = 2
    ElseIf InStr(t, g_f(4)) > 0 Then
        WishBranch = 4
    ElseIf InStr(t, g_f(3)) > 0 Then
        WishBranch = 3
    ElseIf InStr(t, g_f(1)) > 0 Then
        WishBranch = 1
    End If
End Function

Private Function IsRepeater(ByVal x As Variant) As Boolean
    If IsError(x) Then Exit Function
    IsRepeater = (Norm(x) = g_yes)
End Function

Private Function BirthYear(ByVal x As Variant, Optional ByVal date1904 As Boolean = False) As Long
    Dim d As Date
    If TryBirthDate(x, date1904, d) Then BirthYear = Year(d)
End Function

Private Function DecText(ByVal d As Long) As Variant
    Select Case d
    Case 1 To 4: DecText = g_br(d)
    Case D_ORIENT: DecText = L("orient")
    Case D_REPEAT: DecText = L("repeat")
    Case Else: DecText = Empty
    End Select
End Function

' ---------- summary sheet ----------

Private Sub ReportColumnWidth(ByVal ws As Worksheet, ByVal col As Long, ByVal widthPixels As Double)
    Dim k As Long
    With ws.Columns(col)
        .ColumnWidth = 20
        For k = 1 To 3
            If .Width > 0 Then .ColumnWidth = .ColumnWidth * (widthPixels * 0.75) / .Width
        Next k
    End With
End Sub

Private Sub FormatReportGrid(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub

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
    Case "rep_country": L = Txt("627 644 62C 645 647 648 631 64A 629 20 627 644 62C 632 627 626 631 64A 629 20 627 644 62F 64A 645 642 631 627 637 64A 629 20 627 644 634 639 628 64A 629")
    Case "rep_ministry": L = Txt("648 632 627 631 629 20 627 644 62A 631 628 64A 629 20 627 644 648 637 646 64A 629")
    Case "rep_directorate": L = Txt("645 62F 64A 631 64A 629 20 627 644 62A 631 628 64A 629 20 644 648 644 627 64A 629 20 63A 631 62F 627 64A 629")
    Case "rep_school": L = Txt("62B 627 646 648 64A 629 20 62D 64A 20 627 644 634 64A 62E 20 639 627 645 631 20 28 628 631 64A 627 646 29")
    Case "rep_title": L = Txt("62D 635 64A 644 629 20 627 644 62A 648 62C 64A 647 20 627 644 62A 62F 631 64A 62C 64A 20 627 644 62C 630 639 20 627 644 645 634 62A 631 643 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627 20 62D 633 628 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A")
    Case "rep_subtitle": L = Txt("627 642 62A 631 627 62D 627 62A 20 627 644 62A 648 62C 64A 647 20 644 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646 20 62D 633 628 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A")
    Case "rep_registered": L = Txt("639 62F 62F") & vbLf & Txt("627 644 645 633 62C 644 64A 646")
    Case "rep_passed": L = Txt("639 62F 62F 20 627 644 646 627 62C 62D 64A 646")
    Case "rep_repeated": L = Txt("627 644 645 642 62A 631 62D 64A 646 20 644 644 625 639 627 62F 629")
    Case "rep_vocational": L = Txt("627 644 645 642 62A 631 62D 64A 646 20 644 644 62A 648 62C 64A 647") & vbLf & Txt("644 644 62A 643 648 64A 646 20 627 644 645 647 646 64A")
    Case "rep_pct": L = Txt("627 644 646 633 628 629 20 627 644 645 626 648 64A 629")
    Case "rep_wish1": L = Txt("627 644 631 63A 628 629 20 627 644 623 648 644 649")
    Case "rep_decisions": L = Txt("627 642 62A 631 627 62D 20 627 644 62A 648 62C 64A 647")
    Case "rep_met": L = Txt("62A 644 628 64A 629 20 627 644 631 63A 628 629 20 627 644 623 648 644 649")
    Case "f_math": L = Txt("631 64A 627 636")
    Case "f_tech": L = Txt("62A 642 646 64A")
    Case "f_mgmt": L = Txt("62A 633 64A 64A 631")
    Case "f_exp": L = Txt("639 644 648 645")
    End Select
End Function

Private Function OpenWb(ByVal fileName As String, ByVal asReadOnly As Boolean, ByRef opened As Boolean) As Workbook
    Dim wb As Workbook, fullPath As String
    opened = False
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 1, , "Save the tool workbook before running this procedure."
    End If
    fullPath = ThisWorkbook.Path & Application.PathSeparator & fileName
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

' Normalizes Arabic text: alef/yaa/taa-marbuta variants, tatweel, NBSP, Arabic-Indic digits, extra spaces
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

Private Function IsNum(ByVal v As Variant) As Boolean
    If VarType(v) <> vbDouble Then Exit Function
    IsNum = (CDbl(v) >= 0 And CDbl(v) <= 20)
End Function

Private Function Rnd2(ByVal x As Double) As Double
    Rnd2 = Application.WorksheetFunction.Round(x, 2)
End Function

' Decimal multiplication avoids flooring a binary artifact such as 28.999999999.
Private Function ShareCount(ByVal share As Double, ByVal count As Long) As Long
    If share < 0 Or share > 1 Or count < 0 Then
        Err.Raise vbObjectError + 55, , "Invalid percentage or student count."
    End If
    ShareCount = CLng(Fix(CDec(share) * CDec(count)))
End Function

' Ceiling for experimental sciences and management; decimal arithmetic avoids drift.
Private Function ShareCountUp(ByVal share As Double, ByVal count As Long) As Long
    If share < 0 Or share > 1 Or count < 0 Then
        Err.Raise vbObjectError + 55, , "Invalid percentage or student count."
    End If
    ShareCountUp = CLng(-Int(-CDec(share) * CDec(count)))
End Function

' Top 5% use nearest integer, with an exact half rounded upward.
Private Function ShareCountNearest(ByVal share As Double, ByVal count As Long) As Long
    If share < 0 Or share > 1 Or count < 0 Then
        Err.Raise vbObjectError + 55, , "Invalid percentage or student count."
    End If
    ShareCountNearest = CLng(Fix(CDec(share) * CDec(count) + CDec(0.5)))
End Function

' Lock everything except wishes (O:P), then protect the sheet (no password, filtering allowed)

Private Sub ResetStats()
    Erase m_sWish
    Erase m_sDec
    Erase m_sMet
    m_nStud = 0: m_nPass = 0: m_nRep = 0: m_nOri = 0
End Sub

Private Sub ValidateGuidanceData(ByRef v As Variant, ByVal date1904 As Boolean, ByVal firstRow As Long)
    Dim i As Long, c As Long, b1 As Long, b2 As Long, d As Date, location As String
    Dim problems As Object
    Set problems = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(v, 1)
        location = "Data!" & (firstRow + i - 1)
        If IsError(v(i, 2)) Or BlankValue(v(i, 2)) Then AddValidationIssue problems, location, "B"
        For c = 6 To 13
            If Not IsNum(v(i, c)) Then AddValidationIssue problems, location, Chr$(64 + c)
        Next c
        For c = 17 To 20
            If Not IsNum(v(i, c)) Then AddValidationIssue problems, location, Chr$(64 + c)
        Next c
        If Not BlankValue(v(i, C_BIRTH)) Then
            If Not TryBirthDate(v(i, C_BIRTH), date1904, d) Then AddValidationIssue problems, location, "C"
        End If
        b1 = WishBranch(v(i, C_W1)): b2 = WishBranch(v(i, C_W2))
        If IsError(v(i, C_W1)) Or (Not BlankValue(v(i, C_W1)) And b1 = 0) Then
            AddValidationIssue problems, location, "O"
        End If
        If IsError(v(i, C_W2)) Or (Not BlankValue(v(i, C_W2)) And b2 = 0) Then
            AddValidationIssue problems, location, "P"
        End If
        ' Empty + empty remains valid.
        If b1 > 0 And b1 = b2 Then AddValidationIssue problems, location, "O=P"
    Next i
    RaiseValidationIssues problems, "Letters=invalid/missing columns; O=P=repeated non-empty wishes."
End Sub

' Leftovers never fill mathematics or technical mathematics without a request.
' Select the higher eligible profile of experimental sciences and management.
' Experimental sciences wins a tied profile when a seat remains.
Private Function BestRemaining(ByVal i As Long, ByRef counts() As Long, ByRef capacities() As Long) As Long
    Dim b As Long, best As Double
    best = -1
    For b = 3 To 4      ' leftovers: only experimental sciences or management
        If counts(b) < capacities(b) Then
            If Eligible(i, b, 2) Then
                If m_prof(i, b) > best + EPS Then
                    best = m_prof(i, b)
                    BestRemaining = b
                End If
            End If
        End If
    Next b
End Function

' Workbook-aware birth date parsing. Invalid dates never become missing dates.
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

' Collect every invalid row before writing. The dialog lists the first 20.
' Short column codes keep the entire summary within the MsgBox text budget.
Private Sub AddValidationIssue(ByVal problems As Object, ByVal location As String, ByVal detail As String)
    If problems.Exists(location) Then
        If InStr(1, " " & problems(location) & " ", " " & detail & " ", vbBinaryCompare) = 0 Then
            problems(location) = problems(location) & " " & detail
        End If
    Else
        problems.Add location, detail
    End If
End Sub

Private Sub RaiseValidationIssues(ByVal problems As Object, ByVal legend As String)
    Const MAX_ROWS As Long = 20
    Const MAX_LINE As Long = 40
    Dim key As Variant, shown As Long, line As String, msg As String
    If problems.Count = 0 Then Exit Sub
    msg = "Invalid rows: " & problems.Count & " (up to 20 shown)." & vbLf & legend
    For Each key In problems.Keys
        If shown >= MAX_ROWS Then Exit For
        line = CStr(key) & ": " & problems(key)
        If Len(line) > MAX_LINE Then line = Left$(line, MAX_LINE - 3) & "..."
        msg = msg & vbLf & line
        shown = shown + 1
    Next key
    msg = msg & vbLf & "No data written."
    Err.Raise vbObjectError + 58, , msg
End Sub

' Same 15-row report, with two branches and four physical columns.

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
    Case "header_expected": ArtsText = Txt("627 644 639 646 648 627 646 20 627 644 645 62A 648 642 639 20 64A 62A 636 645 646 3A 20")
    Case "h_number": ArtsText = Txt("631 642 645")
    Case "h_name": ArtsText = Txt("627 633 645")
    Case "h_birth": ArtsText = Txt("645 64A 644 627 62F")
    Case "h_sex": ArtsText = Txt("62C 646 633")
    Case "h_repeat": ArtsText = Txt("625 639 627 62F 629")
    Case "h_arabic": ArtsText = Txt("639 631 628 64A")
    Case "h_math": ArtsText = Txt("631 64A 627 636 64A 627 62A")
    Case "h_physics": ArtsText = Txt("641 64A 632 64A 627")
    Case "h_natural": ArtsText = Txt("637 628 64A 639")
    Case "h_history": ArtsText = Txt("62A 627 631 64A 62E")
    Case "h_french": ArtsText = Txt("641 631 646 633 64A")
    Case "h_english": ArtsText = Txt("627 646 62C 644 64A 632")
    Case "h_info": ArtsText = Txt("645 639 644 648 645 627 62A")
    Case "h_info_alt": ArtsText = Txt("625 639 644 627 645")
    Case "h_technology": ArtsText = Txt("62A 643 646 648 644 648 62C")
    Case "h_average": ArtsText = Txt("645 639 62F 644 20 627 644 641 635 644")
    Case "h_avg_word": ArtsText = Txt("645 639 62F 644")
    Case "h_term_word": ArtsText = Txt("641 635 644")
    Case "general_report": ArtsText = Txt("62D 648 635 644 629 20 639 627 645 629")
    Case "general_title": ArtsText = Txt("62D 635 64A 644 629 20 627 644 62A 648 62C 64A 647 20 627 644 646 647 627 626 64A")
    Case "general_subtitle": ArtsText = Txt("627 642 62A 631 627 62D 627 62A 20 627 644 62A 648 62C 64A 647 20 644 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646")
    Case "track_science": ArtsText = Txt("62C 20 645 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627")
    Case "track_arts": ArtsText = Txt("62C 20 645 20 622 62F 627 628")
    Case "male": ArtsText = Txt("630 643 631")
    Case "female": ArtsText = Txt("623 646 62B 649")
    End Select
End Function

' Find the table by its A2 anchor, preserving an existing table name.
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
    Dim firstWish As Long, decisionCol As Long
    If ws.Name = "Data01" Then firstWish = 11 Else firstWish = 15
    If ws.Name = "Data01" Then decisionCol = 15 Else decisionCol = 21
    ws.Cells.Locked = True
    If Not body Is Nothing Then
        body.Columns(firstWish).Resize(, 2).Locked = False
        body.Columns(decisionCol).Locked = False
    End If
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
    ' Keep every report formula and its layout; only old decisions are invalidated.
    lastRow = ws.Cells(ws.Rows.Count, decisionCol).End(xlUp).Row
    If lastRow >= 3 Then ws.Range(ws.Cells(3, decisionCol), ws.Cells(lastRow, decisionCol)).ClearContents
    If Not report Is Nothing Then report.Calculate
End Sub

Private Function ColumnLetter(ByVal n As Long) As String
    Do While n > 0
        ColumnLetter = Chr$(65 + (n - 1) Mod 26) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function


' A placeholder row, whitespace, or formulas returning empty strings is not a pupil.
' A numeric zero, an error, or any other data makes the table non-empty and validated.
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

Private Sub BuildGeneralReport(ByVal wb As Workbook)
    Const COUNT_FORMAT As String = "00"
    Const PERCENT_FORMAT As String = "[=1]%0;%00.00"
    Dim ws As Worksheet, reportTable As ListObject
    Dim widths As Variant, r As Long, c As Long, b As Long
    Dim science As ListObject, arts As ListObject
    Dim oldReport As Worksheet, nm As Variant
    Set science = ReportSourceTable(wb, "Data", 21)
    Set arts = ReportSourceTable(wb, "Data01", 15)
    Set ws = SummarySheet(wb, False, True)
    ws.Unprotect Password:=""
    If ws.ProtectContents Then Err.Raise vbObjectError + 59, , "Unprotect the general summary first."
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Unlist
    Loop
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.DisplayRightToLeft = True
    ws.Rows("1:18").Hidden = False
    ws.Columns("A:F").Hidden = False
    widths = Array(173, 130, 148, 150, 173, 161)
    For c = 1 To 6
        ReportColumnWidth ws, c, CDbl(widths(c - 1))
    Next c
    ws.Rows(1).RowHeight = 34.5
    ws.Rows(2).RowHeight = 27.75
    ws.Rows(3).RowHeight = 26.25
    ws.Rows(4).RowHeight = 27.75
    ws.Rows(5).RowHeight = 26.25
    ws.Rows(6).RowHeight = 57.75
    ws.Rows(7).RowHeight = 29.25
    ws.Rows("8:9").RowHeight = 33
    ws.Rows(10).RowHeight = 36.75
    ws.Rows(11).RowHeight = 34.5
    ws.Rows("12:18").RowHeight = 25.5
    With ws.Range("A1:F18")
        .Font.Name = "Dubai"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .Interior.Pattern = xlNone
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .ShrinkToFit = False
        .NumberFormat = "General"
    End With
    For r = 1 To 5
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Merge
    Next r
    ws.Range("A10:F10").Merge
    ws.Range("A1").Value2 = L("rep_country")
    ws.Range("A2").Value2 = L("rep_ministry")
    ws.Range("A3").Value2 = L("rep_directorate")
    ws.Range("A4").Value2 = L("rep_school")
    ws.Range("A5").Value2 = ArtsText("general_title")
    ws.Range("A10").Value2 = ArtsText("general_subtitle")

    ws.Range("B6").Value2 = L("rep_registered")
    ws.Range("C6").Value2 = L("rep_pct")
    ws.Range("D6").Value2 = L("rep_passed")
    ws.Range("E6").Value2 = L("rep_repeated")
    ws.Range("F6").Value2 = L("rep_vocational")
    ws.Range("B6:F6").WrapText = True
    ws.Range("A7").Value2 = ArtsText("track_science")
    ws.Range("A8").Value2 = ArtsText("track_arts")
    ws.Range("A9").Value2 = L("total")
    ws.Range("B7").Formula = "=" & ReportCount(science, "registered")
    ws.Range("D7").Formula = "=" & ReportCount(science, "passed")
    ws.Range("E7").Formula = "=" & ReportCount(science, "decision", L("repeat"))
    ws.Range("F7").Formula = "=" & ReportCount(science, "decision", L("orient"))
    ws.Range("B8").Formula = "=" & ReportCount(arts, "registered")
    ws.Range("D8").Formula = "=" & ReportCount(arts, "passed")
    ws.Range("E8").Formula = "=" & ReportCount(arts, "decision", L("repeat"))
    ws.Range("F8").Formula = "=" & ReportCount(arts, "decision", L("orient"))
    For c = 2 To 6
        If c <> 3 Then ws.Cells(9, c).FormulaR1C1 = "=SUM(R7C:R8C)"
    Next c
    ws.Range("C7:C9").FormulaR1C1 = "=IF(R9C2=0,0,RC2/R9C2)"

    ws.Range("A11").Value2 = L("c_branch")
    ws.Range("B11").Value2 = L("rep_wish1")
    ws.Range("C11").Value2 = L("rep_decisions")
    ws.Range("D11").Value2 = L("rep_pct")
    ws.Range("E11").Value2 = L("rep_met")
    ws.Range("F11").Value2 = L("rep_pct") & "2"
    For b = 1 To 4
        r = 11 + b
        ws.Cells(r, 1).Value2 = g_br(b)
        ws.Cells(r, 2).Formula = "=" & ReportCount(science, "wish", g_br(b))
        ws.Cells(r, 3).Formula = "=" & ReportCount(science, "decision", g_br(b))
        ws.Cells(r, 5).Formula = "=" & ReportCount(science, "met", g_br(b))
    Next b
    ' The reference lists languages first, then arts/philosophy.
    ws.Range("A16").Value2 = ArtsText("lang")
    ws.Range("B16").Formula = "=" & ReportCount(arts, "wish", ArtsText("lang"))
    ws.Range("C16").Formula = "=" & ReportCount(arts, "decision", ArtsText("lang"))
    ws.Range("E16").Formula = "=" & ReportCount(arts, "met", ArtsText("lang"))
    ws.Range("A17").Value2 = ArtsText("arts")
    ws.Range("B17").Formula = "=" & ReportCount(arts, "wish", ArtsText("arts"))
    ws.Range("C17").Formula = "=" & ReportCount(arts, "decision", ArtsText("arts"))
    ws.Range("E17").Formula = "=" & ReportCount(arts, "met", ArtsText("arts"))
    ws.Range("A18").Value2 = L("total")
    ws.Range("B18").Formula = "=SUM(B12:B17)"
    ws.Range("C18").Formula = "=SUM(C12:C17)"
    ws.Range("E18").Formula = "=SUM(E12:E17)"
    ' Proposal shares are calculated within their own track: science and arts separately.
    ws.Range("D12:D15").FormulaR1C1 = "=IF(SUM(R12C3:R15C3)=0,0,RC3/SUM(R12C3:R15C3))"
    ws.Range("D16:D17").FormulaR1C1 = "=IF(SUM(R16C3:R17C3)=0,0,RC3/SUM(R16C3:R17C3))"
    ws.Range("D18").Formula = "=IF(C18=0,0,C18/C18)"

    ' Wish satisfaction is measured against the decisions for that same branch.
    ws.Range("F12:F17").FormulaR1C1 = "=IF(RC3=0,0,RC5/RC3)"
    ws.Range("F18").Formula = "=IF(C18=0,0,E18/C18)"

    Set reportTable = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("A11:F18"), XlListObjectHasHeaders:=xlYes)
    reportTable.TableStyle = ""
    reportTable.ShowTotals = False
    reportTable.ShowTableStyleFirstColumn = False
    reportTable.ShowTableStyleLastColumn = False
    reportTable.ShowTableStyleRowStripes = False
    reportTable.ShowTableStyleColumnStripes = False
    reportTable.ShowAutoFilter = True
    FormatReportGrid ws.Range("A6:F9")
    FormatReportGrid ws.Range("A11:F18")
    ws.Range("B8:F9").Interior.Color = RGB(237, 237, 237)
    ws.Range("B7:F9").NumberFormat = COUNT_FORMAT
    ws.Range("C7:C9").NumberFormat = PERCENT_FORMAT
    ws.Range("B12:C18").NumberFormat = COUNT_FORMAT
    ws.Range("E12:E18").NumberFormat = COUNT_FORMAT
    ws.Range("D12:D18").NumberFormat = PERCENT_FORMAT
    ws.Range("F12:F18").NumberFormat = PERCENT_FORMAT
    ws.Range("B7:F8").Font.Bold = False
    ws.Range("B12:F17").Font.Bold = False
    ws.Range("A9:F9").Font.Bold = True
    ws.Range("A18:F18").Font.Bold = True
    ws.Calculate

    ' Remove only the superseded report tabs, after the combined report is complete.
    For Each nm In Array(ArtsText("science_report"), ArtsText("arts_report"), ArtsText("legacy_report"))
        Set oldReport = FindWorksheet(wb, CStr(nm))
        If Not oldReport Is Nothing Then oldReport.Delete
    Next nm
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$F$18"
    On Error GoTo 0
End Sub

Private Function ReportSourceTable(ByVal wb As Workbook, ByVal sheetName As String, ByVal width As Long) As ListObject
    Dim ws As Worksheet, lo As ListObject
    Set ws = FindWorksheet(wb, sheetName)
    If ws Is Nothing Then Exit Function
    Set lo = GuidanceTable(ws, width)
    If lo.DataBodyRange Is Nothing Then
        ws.Unprotect Password:=""
        If ws.ProtectContents Then Err.Raise vbObjectError + 81, , "Unprotect " & ws.Name & " before running."
        lo.ListRows.Add
        LockAndProtect ws, lo.DataBodyRange
    End If
    Set ReportSourceTable = lo
End Function

Private Function FormulaText(ByVal value As String) As String
    FormulaText = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function FormulaNumber(ByVal value As Double) As String
    FormulaNumber = Replace(CStr(value), Application.International(xlDecimalSeparator), ".")
End Function

Private Function ReportColumnRef(ByVal lo As ListObject, ByVal columnNo As Long) As String
    Dim header As String
    header = lo.ListColumns(columnNo).Name
    header = Replace(header, "'", "''")
    header = Replace(header, "[", "'[")
    header = Replace(header, "]", "']")
    header = Replace(header, "#", "'#")
    header = Replace(header, "@", "'@")
    ReportColumnRef = lo.Name & "[" & header & "]"
End Function

Private Function ReportPatterns(ByVal label As String) As Variant
    Dim pattern As String
    pattern = "*" & Replace(label, " ", "*") & "*"
    ReportPatterns = Array(pattern)
End Function

Private Function ReportCount(ByVal lo As ListObject, ByVal mode As String, _
                             Optional ByVal label As String = "", Optional ByVal sex As Long = 0) As String
    Dim criteria As String, expression As String, term As String
    Dim decisionCol As Long, wishCol As Long, averageCol As Long, sexLabel As String
    Dim sexes As Variant, decisions As Variant, wishes As Variant, sx As Variant, dc As Variant, wi As Variant
    If lo Is Nothing Then
        ReportCount = "0"
        Exit Function
    End If
    If lo.Parent.Name = "Data01" Then
        decisionCol = 15: wishCol = 11: averageCol = 10
    Else
        decisionCol = 21: wishCol = 15: averageCol = 13
    End If
    criteria = ReportColumnRef(lo, 2) & "," & FormulaText("?*")
    sexes = Array(""): decisions = Array(""): wishes = Array("")
    If sex > 0 Then
        If sex = 1 Then sexLabel = ArtsText("male") Else sexLabel = ArtsText("female")
        sexes = ReportPatterns(sexLabel)
    End If
    If mode = "passed" Then
        criteria = criteria & "," & ReportColumnRef(lo, averageCol) & "," & FormulaText(">=") & "&" & FormulaNumber(PASS_MARK - EPS)
    ElseIf mode = "decision" Or mode = "met" Then
        decisions = ReportPatterns(label)
    End If
    If mode = "wish" Or mode = "met" Then wishes = ReportPatterns(label)
    For Each sx In sexes
        For Each dc In decisions
            For Each wi In wishes
                term = criteria
                If Len(CStr(sx)) > 0 Then term = term & "," & ReportColumnRef(lo, 4) & "," & FormulaText(CStr(sx))
                If Len(CStr(dc)) > 0 Then term = term & "," & ReportColumnRef(lo, decisionCol) & "," & FormulaText(CStr(dc))
                If Len(CStr(wi)) > 0 Then term = term & "," & ReportColumnRef(lo, wishCol) & "," & FormulaText(CStr(wi))
                If Len(expression) > 0 Then expression = expression & "+"
                expression = expression & "COUNTIFS(" & term & ")"
            Next wi
        Next dc
    Next sx
    ReportCount = "(" & expression & ")"
End Function

Public Sub RefreshReports02()
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, nm As Variant
    Dim opened As Boolean, found As Boolean, failed As Boolean, touched As Boolean
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldEvents As Boolean, stateSaved As Boolean
    Dim errText As String, width As Long
    On Error GoTo EH
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    stateSaved = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    InitText
    Set wb = OpenWb(DST_FILE, False, opened)
    If wb.ReadOnly Then Err.Raise vbObjectError + 2, , DST_FILE & " is open read-only."
    ' Validate the layout of every available track before making any changes.
    For Each nm In Array("Data", "Data01")
        Set ws = FindWorksheet(wb, CStr(nm))
        If Not ws Is Nothing Then
            found = True
            If ws.Name = "Data01" Then width = 15 Else width = 21
            Set lo = GuidanceTable(ws, width)
        End If
    Next nm
    If Not found Then Err.Raise vbObjectError + 84, , "Neither Data nor Data01 exists in " & DST_FILE & "."
    For Each nm In Array("Data", "Data01")
        Set ws = FindWorksheet(wb, CStr(nm))
        If Not ws Is Nothing Then
            touched = True
            CheckTrackWrite wb, ws
            If ws.Name = "Data01" Then width = 15 Else width = 21
            Set lo = GuidanceTable(ws, width)
            LockAndProtect ws, lo.DataBodyRange
        End If
    Next nm
    BuildGeneralReport wb
    Application.Calculation = xlCalculationAutomatic
    wb.Save
Done:
    On Error Resume Next
    If failed And opened Then
        wb.Close SaveChanges:=False
    ElseIf failed And touched And Not wb Is Nothing Then
        For Each nm In Array("Data", "Data01")
            Set ws = FindWorksheet(wb, CStr(nm))
            If Not ws Is Nothing Then
                ws.Unprotect Password:=""
                If ws.Name = "Data01" Then width = 15 Else width = 21
                Set lo = GuidanceTable(ws, width)
                LockAndProtect ws, lo.DataBodyRange
            End If
        Next nm
    End If
    If stateSaved Then
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
    End If
    On Error GoTo 0
    If failed Then
        MsgBox "Report refresh failed:" & vbLf & errText, vbCritical, "Final reports"
    Else
        MsgBox "The live general report was updated from the existing Data and Data01 values.", vbInformation, "Final reports"
    End If
    Exit Sub
EH:
    failed = True
    errText = Err.Description
    If touched And Not opened Then errText = errText & vbLf & vbLf & ArtsText("failure_unsaved")
    Resume Done
End Sub

Public Function Orientation02Succeeded() As Boolean
    Orientation02Succeeded = m_runSucceeded
End Function
