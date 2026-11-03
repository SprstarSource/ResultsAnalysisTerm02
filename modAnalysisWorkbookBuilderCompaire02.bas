Attribute VB_Name = "modBiuldAnalysisfilesCompare02"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesCompare02
' Purpose : Adds the Compare sheet: term 1 against term 2, for the data of the FIRST
'           table of the Analysis sheets only (per subject + the "passed" cells).
' Design  : - One header for rows and columns (the subjects), then every item in two rows:
'               <item>   term 1   values
'                        term 2   values   (directly under term 1)
'           - The values are plain links to Analysis01 / Analysis02 (same cell addresses,
'             because both sheets have the same layout): nothing is calculated twice.
'           - Two extra figures per section, calculated from Data:
'               pupils who improved   : average 2 - average 1 >= 1.00
'               pupils who declined   : average 1 - average 2 >= 1.00
'             (the difference is rounded to 2 decimals so 1.00 exactly always counts).
'           - Same colours as the Analysis sheets for averages and success rates.
'==============================================================================

Public Function baf_BuildCompareSheet(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet, b As Long, lastCol As Long, n As Long, col As Long, k As Long, i As Long
    Dim r0 As Long, rHdr As Long, rM As Long, rP As Long, nQ As Long
    Dim sc As String, secN As String, a1 As String, a2 As String, ttl As String
    Dim s1 As String, s2 As String, tag1 As String, tag2 As String
    Dim hdrText() As String, fm() As Variant, cell1 As String
    Dim lblM As Variant, fmtM As Variant
    Dim nextRow As Long, lastRowUsed As Long

    baf_BuildCompareSheet = False
    errText = ""
    On Error GoTo EH

    If Not baf_DiscoverData(wb, errText) Then Exit Function
    If baf_gBlkN < 1 Then errText = "The Analysis sheets must be built before the Compare sheet": Exit Function

    n = baf_gSubjN
    lastCol = baf_FIRST_NUM_COL + n
    sc = baf_ColRef(baf_gHdr(baf_gSecCol))
    a1 = baf_ColRef(baf_gHdr(baf_gAvgC1))
    a2 = baf_ColRef(baf_gHdr(baf_gAvgC2))
    s1 = baf_ANALYSIS_PREFIX & "01"
    s2 = baf_ANALYSIS_PREFIX & "02"
    tag1 = baf_ArT("641") & "1"
    tag2 = baf_ArT("641") & "2"

    ReDim hdrText(baf_FIRST_NUM_COL To lastCol)
    For k = 1 To n
        hdrText(baf_FIRST_NUM_COL + k - 1) = baf_gSubjName(k)
    Next k
    hdrText(lastCol) = baf_ArT("627 644 645 639 62F 644 20 627 644 641 635 644 64A")

    ' the 7 items of the first table, in the order of the Analysis sheets (offset from the header row)
    lblM = Array(baf_ArT("645 639 62F 644 627 62A 20 627 644 645 648 627 62F"), baf_ArT("639 62F 62F 20 627 644 62A 644 645 64A 630 627 62A 20 627 644 646 627 62C 62D 627 62A"), baf_ArT("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646 20 28 630 643 648 631 29"), _
                 baf_ArT("646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 643 644 20 645 627 62F 629"), baf_ArT("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 645 639 64A 62F 64A 646 20 627 644 646 627 62C 62D 64A 646"), baf_ArT("623 639 644 649 20 645 639 62F 644"), baf_ArT("623 62F 646 649 20 645 639 62F 644"))
    fmtM = Array(baf_FMT_DEC, baf_FMT_COUNT, baf_FMT_COUNT, baf_FMT_PCT, baf_FMT_COUNT, baf_FMT_DEC, baf_FMT_DEC)

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = baf_COMPARE_SHEET
    ws.DisplayRightToLeft = True

    nextRow = 1
    For b = 1 To baf_gBlkN
        r0 = nextRow
        secN = CStr(baf_gBlkSec(b))
        nQ = baf_gBlkNQ(b)
        rHdr = r0 + 4
        rM = rHdr + 1
        rP = rM + 14 + 1
        lastRowUsed = rP + 2 * nQ - 1

        ' ---- banner
        ttl = baf_ArT("645 642 627 631 646 629 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644 20 648 627 644 641 635 644 20 627 644 62B 627 646 64A") & " " & baf_ArT("642 633 645 20") & baf_gClassTitle
        If baf_gSecN > 1 Then ttl = ttl & " " & Format$(baf_gBlkSec(b), "00")
        baf_MergeText ws, r0, 1, lastCol, ttl
        baf_StyleBanner ws.Range(ws.Cells(r0, 1), ws.Cells(r0, lastCol))
        ws.Rows(r0).RowHeight = baf_TITLE_ROW_HEIGHT

        ' ---- improved / declined (average difference of 1.00 or more)
        baf_PutStrip ws, r0 + 1, 1, 3, 4, baf_ArT("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 645 62A 641 648 642 64A 646 20 28 641 627 631 642 20 2B 31 2E 30 30 20 641 623 643 62B 631 29"), _
                 "=SUMPRODUCT((" & sc & "=" & secN & ")*(" & a1 & ">0)*(" & a2 & ">0)*(ROUND(" & a2 & "-" & a1 & ",2)>=1))"
        baf_PutStrip ws, r0 + 2, 1, 3, 4, baf_ArT("639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 645 62A 631 627 62C 639 64A 646 20 28 641 627 631 642 20 2D 31 2E 30 30 20 641 623 643 62B 631 29"), _
                 "=SUMPRODUCT((" & sc & "=" & secN & ")*(" & a1 & ">0)*(" & a2 & ">0)*(ROUND(" & a1 & "-" & a2 & ",2)>=1))"
        With ws.Cells(r0 + 1, baf_FIRST_NUM_COL)
            .Interior.Color = RGB(0, 128, 64): .Font.Color = RGB(255, 255, 255)
        End With
        With ws.Cells(r0 + 2, baf_FIRST_NUM_COL)
            .Interior.Color = RGB(192, 0, 0): .Font.Color = RGB(255, 255, 255)
        End With
        ws.Rows(r0 + 3).RowHeight = 6

        ' ---- one header (subjects) for the whole block
        baf_PutHeaderRow ws, rHdr, lastCol, baf_ArT("627 644 645 648 627 62F"), hdrText

        ' ---- the 7 items : term 1 row, then term 2 row right under it
        ReDim fm(1 To 14, 1 To lastCol - baf_FIRST_NUM_COL + 1)
        For i = 0 To 6
            For col = baf_FIRST_NUM_COL To lastCol
                cell1 = ws.Cells(1, col).Address(False, False)
                cell1 = Left$(cell1, Len(cell1) - 1) & CStr(baf_gBlkH1(b) + 1 + i)
                fm(2 * i + 1, col - baf_FIRST_NUM_COL + 1) = "=" & s1 & "!" & cell1
                fm(2 * i + 2, col - baf_FIRST_NUM_COL + 1) = "=" & s2 & "!" & cell1
            Next col
        Next i
        ws.Range(ws.Cells(rM, baf_FIRST_NUM_COL), ws.Cells(rM + 13, lastCol)).Formula = fm
        baf_StyleGrid ws.Range(ws.Cells(rM, baf_FIRST_NUM_COL), ws.Cells(rM + 13, lastCol))
        For i = 0 To 6
            ws.Range(ws.Cells(rM + 2 * i, baf_FIRST_NUM_COL), ws.Cells(rM + 2 * i + 1, lastCol)).NumberFormat = CStr(fmtM(i))
            PutCompareLabel ws, rM + 2 * i, CStr(lblM(i)), tag1, tag2
            ' term 2 row: very light amber background so each pair is easy to read
            ws.Range(ws.Cells(rM + 2 * i + 1, baf_FIRST_NUM_COL), ws.Cells(rM + 2 * i + 1, lastCol)).Interior.Color = RGB(255, 250, 235)
        Next i
        ' averages (rows 1-2) and success rates (rows 7-8): red below the limit, green from the limit up
        For i = 0 To 1
            baf_ApplyPassColors ws.Range(ws.Cells(rM + i, baf_FIRST_NUM_COL), ws.Cells(rM + i, lastCol)), "=10", "=20"
            baf_ApplyPassColors ws.Range(ws.Cells(rM + 6 + i, baf_FIRST_NUM_COL), ws.Cells(rM + 6 + i, lastCol)), "=0.5", "=1"
        Next i
        ws.Rows(rM + 14).RowHeight = 6

        ' ---- the "passed" cells (term 1 row, term 2 row)
        For i = 1 To nQ
            cell1 = ws.Cells(1, baf_FIRST_NUM_COL).Address(False, False)
            cell1 = Left$(cell1, Len(cell1) - 1) & CStr(baf_gBlkQ1(b) + i - 1)
            ws.Cells(rP + 2 * (i - 1), baf_FIRST_NUM_COL).Formula = "=" & s1 & "!" & cell1
            ws.Cells(rP + 2 * (i - 1) + 1, baf_FIRST_NUM_COL).Formula = "=" & s2 & "!" & cell1
            baf_StyleGrid ws.Range(ws.Cells(rP + 2 * (i - 1), baf_FIRST_NUM_COL), ws.Cells(rP + 2 * (i - 1) + 1, baf_FIRST_NUM_COL))
            ws.Range(ws.Cells(rP + 2 * (i - 1), baf_FIRST_NUM_COL), ws.Cells(rP + 2 * (i - 1) + 1, baf_FIRST_NUM_COL)).NumberFormat = baf_FMT_COUNT
            PutCompareLabel ws, rP + 2 * (i - 1), baf_gBlkQLbl(b, i), tag1, tag2
            ws.Cells(rP + 2 * (i - 1) + 1, baf_FIRST_NUM_COL).Interior.Color = RGB(255, 250, 235)
        Next i

        nextRow = lastRowUsed + 3
    Next b

    baf_SetColumnWidths ws, lastCol
    baf_SetupPrint ws, nextRow - 3, lastCol

    baf_BuildCompareSheet = True
    Exit Function
EH:
    errText = "baf_BuildCompareSheet: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function

' item name (merged over two rows) + the term tags "term 1" / "term 2"
Private Sub PutCompareLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, _
                            ByVal tag1 As String, ByVal tag2 As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + 1, 2))
        .Merge
        .Cells(1, 1).Value2 = lbl
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = baf_gFontName: .Font.Size = 13: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ShrinkToFit = False
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    ws.Cells(r, 3).Value2 = tag1
    ws.Cells(r + 1, 3).Value2 = tag2
    With ws.Cells(r, 3)
        .Interior.Color = RGB(221, 235, 247)
    End With
    With ws.Cells(r + 1, 3)
        .Interior.Color = RGB(255, 242, 204)
    End With
    With ws.Range(ws.Cells(r, 3), ws.Cells(r + 1, 3))
        .Font.Name = baf_gFontBase: .Font.Size = 14: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub
