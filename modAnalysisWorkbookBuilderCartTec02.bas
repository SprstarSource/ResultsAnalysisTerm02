Attribute VB_Name = "modBiuldAnalysisfilesCartTec02"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesCartTec02   (Build_CartTec)
' Purpose : Adds the Cart_Tec sheet (one pupil card) to a workbook that already
'           has the Data sheet. Designed for a projector (DataShow).
' Design  : - I1 = record number, linked to a spinner placed over I1.
'           - Every value is INDEX(Data[column], $I$1, 1); nothing else.
'           - The spinner maximum = number of rows of the Data table at build time.
'           - Subjects come from the Data headers (same discovery as the Analysis sheets).
'           - Rank in the section = the first column of Data (the number).
'           - Picture (in-cell) and notes are typed/inserted by hand in Data.
'           - Colours: average row = cyan, then yellow / grey rows one after the other;
'             the difference column is bold (green = up, red = down).
'           - A mark of 0 = not taught / exempt: hidden, except physical education
'             which shows the word "exempt".
'==============================================================================

Private Const FMT_MARK As String = "00.00;-00.00;;@"
Private Const FMT_DIFF As String = "+00.00;-00.00;00.00"
Private Const FMT_HIDE0 As String = "General;-General;;@"
Private Const FMT_DATE0 As String = "yyyy-mm-dd;;;@"
Private Const FMT_RANK0 As String = "00;-00;;@"

Public Function baf_BuildCartTecSheet(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet, sp As Shape
    Dim n As Long, k As Long, r As Long, rLast As Long, rBottom As Long
    Dim accCol As Long, picCol As Long, notesCol As Long
    Dim multi As Boolean, maxV As Long
    Dim rowRef As String, yes As String, fmtExempt As String
    Dim fcs As Object, stopTxt As String, stopSign As String
    Dim clrYel As Long, clrGry As Long, clrCya As Long, clrBan As Long, clrRow As Long
    Dim rI As Range

    baf_BuildCartTecSheet = False
    errText = ""
    On Error GoTo EH

    If Not baf_DiscoverData(wb, errText) Then Exit Function

    clrYel = RGB(255, 255, 102)
    clrGry = RGB(235, 235, 235)
    clrCya = RGB(102, 255, 255)
    clrBan = RGB(255, 255, 0)

    n = baf_gSubjN
    accCol = baf_gSecCol + 2
    picCol = baf_gSecCol + 3
    notesCol = baf_gSecCol + 4
    multi = (baf_gSecN > 1)
    yes = baf_ArT("646 639 645")
    stopSign = ChrW(&HD83D) & ChrW(&HDED1)
    fmtExempt = "00.00;-00.00;" & baf_QT & baf_ArT("625 639 641 640 640 627 621") & baf_QT & ";@"
    rowRef = ",$I$1,1)"
    rLast = 4 + n
    rBottom = rLast
    If rBottom < 14 Then rBottom = 14

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = baf_CART_SHEET
    ws.DisplayRightToLeft = True

    ' ---- widths
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 17
    ws.Columns(3).ColumnWidth = 28
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 0.8
    ws.Columns(6).ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 0.8
    ws.Columns(8).ColumnWidth = 12
    ws.Columns(9).ColumnWidth = 7.1
    ws.Columns(10).ColumnWidth = 4.2
    ws.Columns(11).ColumnWidth = 5.6

    ' ---- record number (linked to the spinner, which is placed over this cell)
    ws.Range("I1").Value = 1
    With ws.Range("I1")
        .Font.Name = baf_gFontNum: .Font.Size = 12: .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .NumberFormat = "0"
    End With

    ' ---- title
    baf_MergeText ws, 1, 1, 8, baf_ArT("628 637 627 642 629 20 641 646 64A 629 20 644 644 62A 644 645 64A 630")
    With ws.Range("A1:H1")
        .Interior.Color = clrBan
        .Font.Name = baf_gFontBase: .Font.Size = 28: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(1).RowHeight = 44

    ' ---- left block : name, birth date, repeat, transition average, rank
    PutCardLabel ws, "A2", baf_ArT("627 644 627 633 645 20 648 627 644 644 642 628"), clrYel
    PutCardValue ws, "A3", "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(2)) & rowRef & "," & baf_QT & baf_QT & ")", FMT_HIDE0
    PutCardLabel ws, "A4", baf_ArT("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F"), clrYel
    PutCardValue ws, "A5", "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(3)) & rowRef & "," & baf_QT & baf_QT & ")", FMT_DATE0
    PutCardLabel ws, "A6", baf_gHdr(5), clrYel
    PutCardValue ws, "A7", "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(5)) & rowRef & "," & baf_QT & baf_QT & ")", "General"
    PutCardLabel ws, "A8", baf_gHdr(accCol), clrYel
    PutCardValue ws, "A9", "=IFERROR(IF(A7=" & baf_QT & yes & baf_QT & "," & baf_QT & baf_QT & ",INDEX(" & baf_ColRef(baf_gHdr(accCol)) & rowRef & ")," & baf_QT & baf_QT & ")", FMT_MARK
    PutCardLabel ws, "A10", baf_ArT("627 644 631 62A 628 629 20 641 64A 20 627 644 642 633 645"), clrYel
    PutCardValue ws, "A11", "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(1)) & rowRef & "," & baf_QT & baf_QT & ")", FMT_RANK0

    ' ---- picture (in-cell picture typed in Data) and notes
    ws.Range("B2:B7").Merge
    If multi Then
        ws.Range("B2").Formula = "=IFERROR(IF($A$12<>" & baf_QT & baf_QT & "," & baf_QT & baf_QT & ",INDEX(" & baf_ColRef(baf_gHdr(picCol)) & rowRef & ")," & baf_QT & baf_QT & ")"
    Else
        ws.Range("B2").Formula = "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(picCol)) & rowRef & "," & baf_QT & baf_QT & ")"
    End If
    With ws.Range("B2:B7")
        .NumberFormat = FMT_HIDE0
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    ws.Range("B8:B11").Merge
    ws.Range("B8").Formula = "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(notesCol)) & rowRef & "," & baf_QT & baf_QT & ")"
    With ws.Range("B8:B11")
        .NumberFormat = FMT_HIDE0
        .Font.Name = baf_gFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With

    ' ---- class title (and, for several sections, the section being browsed: typed by hand)
    If multi Then
        ws.Range("C2").Value2 = baf_gClassTitle
        With ws.Range("C2")
            .Interior.Color = clrYel
            .Font.Name = baf_gFontBase: .Font.Size = 20: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .ShrinkToFit = True
            .ReadingOrder = xlRTL
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
        With ws.Range("C3")
            .NumberFormat = baf_FMT_COUNT
            .Interior.Color = clrGry
            .Font.Name = baf_gFontNum: .Font.Size = 22: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
    Else
        ws.Range("C2:C3").Merge
        ws.Range("C2").Value2 = baf_gClassTitle
        With ws.Range("C2:C3")
            .Interior.Color = clrYel
            .Font.Name = baf_gFontBase: .Font.Size = 20: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
            .ReadingOrder = xlRTL
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
    End If

    ' ---- results block : term 2 | difference | term 1
    ws.Range("D2:D3").Merge
    ws.Range("D2").Value2 = baf_ArT("646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A")
    ws.Range("F2:F3").Merge
    ws.Range("F2").Value2 = baf_ArT("627 644 641 627 631 642")
    ws.Range("H2:H3").Merge
    ws.Range("H2").Value2 = baf_ArT("646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644")
    baf_StyleHead ws.Range("D2:D3")
    baf_StyleHead ws.Range("F2:F3")
    baf_StyleHead ws.Range("H2:H3")
    For Each rI In ws.Range("D2:D3,F2:F3,H2:H3").Areas
        rI.Interior.Color = clrCya
        rI.Font.Color = RGB(0, 0, 0)
        rI.Font.Size = 15
    Next rI
    ws.Range("F2:F3").Font.Size = 18

    ' averages (row 4) and one row per subject
    ws.Range("C4").Value2 = baf_ArT("627 644 645 639 62F 644 20 627 644 641 635 644 64A")
    ws.Range("D4").Formula = "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(baf_gAvgC2)) & rowRef & "," & baf_QT & baf_QT & ")"
    ws.Range("H4").Formula = "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(baf_gAvgC1)) & rowRef & "," & baf_QT & baf_QT & ")"
    ws.Range("F4").Formula = "=IFERROR(D4-H4," & baf_QT & baf_QT & ")"
    For k = 1 To n
        r = 4 + k
        ws.Cells(r, 3).Value2 = baf_gSubjName(k)
        ws.Cells(r, 4).Formula = "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(baf_gSubjC2(k))) & rowRef & "," & baf_QT & baf_QT & ")"
        ws.Cells(r, 8).Formula = "=IFERROR(INDEX(" & baf_ColRef(baf_gHdr(baf_gSubjC1(k))) & rowRef & "," & baf_QT & baf_QT & ")"
        ws.Cells(r, 6).Formula = "=IFERROR(IF(OR(D" & r & "=0,H" & r & "=0)," & baf_QT & baf_QT & ",D" & r & "-H" & r & ")," & baf_QT & baf_QT & ")"
    Next k

    baf_StyleGrid ws.Range(ws.Cells(4, 3), ws.Cells(rLast, 3))
    baf_StyleGrid ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 4))
    baf_StyleGrid ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6))
    baf_StyleGrid ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8))
    With ws.Range(ws.Cells(4, 3), ws.Cells(rLast, 3))
        .Font.Name = baf_gFontName: .Font.Size = 17: .Font.Bold = True
    End With
    ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 4)).Font.Size = 20
    ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8)).Font.Size = 20
    ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6)).Font.Size = 20
    ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6)).Font.Bold = True
    ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 4)).NumberFormat = FMT_MARK
    ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8)).NumberFormat = FMT_MARK
    ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6)).NumberFormat = FMT_DIFF
    If baf_gPEIdx > 0 Then
        ws.Cells(4 + baf_gPEIdx, 4).NumberFormat = fmtExempt
        ws.Cells(4 + baf_gPEIdx, 8).NumberFormat = fmtExempt
    End If

    ' average row = cyan ; subject rows = yellow / grey one after the other
    For Each rI In ws.Range("C4,D4,F4,H4").Areas
        rI.Interior.Color = clrCya
        rI.Font.Bold = True
    Next rI
    For k = 1 To n
        r = 4 + k
        If k Mod 2 = 1 Then clrRow = clrYel Else clrRow = clrGry
        ws.Cells(r, 3).Interior.Color = clrRow
        ws.Cells(r, 4).Interior.Color = clrRow
        ws.Cells(r, 6).Interior.Color = clrRow
        ws.Cells(r, 8).Interior.Color = clrRow
    Next k

    ' ---- same row height everywhere (the picture spans six rows)
    For r = 2 To rBottom
        ws.Rows(r).RowHeight = 32
    Next r

    ' ---- end-of-section message (several sections only)
    '   shows only when C3 (typed by hand) holds a section number and the selected pupil is in another one
    If multi Then
        stopTxt = stopSign & " " & baf_ArT("646 640 647 640 627 64A 640 629 20 627 644 640 642 640 633 640 645") & " " & stopSign
        ws.Range("A12:B14").Merge
        ws.Range("A12").Formula = "=IFERROR(IF($C$3=" & baf_QT & baf_QT & "," & baf_QT & baf_QT & ",IF(INDEX(" & baf_ColRef(baf_gHdr(baf_gSecCol)) & rowRef & _
                                  "<>$C$3," & baf_QT & stopTxt & baf_QT & "," & baf_QT & baf_QT & "))," & baf_QT & baf_QT & ")"
        With ws.Range("A12:B14")
            .Font.Name = baf_gFontBase: .Font.Size = 26: .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = False
            .ShrinkToFit = True
            .ReadingOrder = xlRTL
        End With
    End If

    ' ---- conditional formats
    With ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6))
        Set fcs = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        fcs.Font.Color = RGB(0, 150, 50)
        Set fcs = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        fcs.Font.Color = RGB(226, 54, 0)
    End With
    If multi Then
        ' bright yellow band behind the message
        Set fcs = ws.Range("A12:B14").FormatConditions.Add(Type:=xlExpression, Formula1:="=$A$12<>" & baf_QT & baf_QT)
        fcs.Interior.Color = RGB(255, 255, 0)
        fcs.Font.Color = RGB(255, 0, 0)
        ' hide the data of every other cell while the message is shown
        Set fcs = ws.Range("A3,A5,A7,A9,A11,B8:B11,D4:H" & rLast).FormatConditions.Add( _
                  Type:=xlExpression, Formula1:="=$A$12<>" & baf_QT & baf_QT)
        fcs.NumberFormat = ";;;"
        fcs.Font.Color = RGB(255, 255, 255)
    End If

    ' ---- spinner linked to I1, placed OVER I1 (it hides the number below it)
    maxV = baf_gRows
    If maxV < 1 Then maxV = 1
    If maxV > 30000 Then maxV = 30000
    Set sp = ws.Shapes.AddFormControl(xlSpinner, ws.Range("I1").Left + 2, ws.Range("I1").Top + 2, _
                                      ws.Range("I1").Width - 4, ws.Range("I1").Height + ws.Range("I2").Height - 4)
    sp.Name = "Spinner 1"
    ' the sheet is protected at the end of the build: spinner and I1 must stay usable
    On Error Resume Next
    sp.Locked = False
    ws.Range("I1").Locked = False
    Err.Clear
    On Error GoTo EH
    With sp.ControlFormat
        .LinkedCell = "$I$1"
        .Min = 1
        .Max = maxV
        .SmallChange = 1
        .Value = 1
    End With

    ' ---- window
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    ws.Range("A1").Select
    Err.Clear
    On Error GoTo EH

    baf_BuildCartTecSheet = True
    Exit Function
EH:
    errText = "baf_BuildCartTecSheet: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function

Private Sub PutCardLabel(ByVal ws As Worksheet, ByVal addr As String, ByVal txt As String, ByVal clr As Long)
    ws.Range(addr).Value2 = txt
    With ws.Range(addr)
        .Interior.Color = clr
        .Font.Name = baf_gFontName: .Font.Size = 17: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub PutCardValue(ByVal ws As Worksheet, ByVal addr As String, ByVal f As String, ByVal fmt As String)
    With ws.Range(addr)
        .Formula = f
        .NumberFormat = fmt
        .Font.Name = baf_gFontName: .Font.Size = 17: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub
