Attribute VB_Name = "MxXlsLoSetLc"
Option Explicit
Option Compare Text
Const CNs$ = "Set.Lc"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLoSetLc."

Sub SetLccWdt(L As ListObject, CC$, W):           Dim C: For Each C In Split(CC): SetLcWdt L, C, W:   Next: End Sub
Sub SetLccFmt(L As ListObject, CC$, Fmt$):        Dim C: For Each C In Split(CC): SetLcFmt L, C, Fmt: Next: End Sub
Sub SetLccAsSum(L As ListObject, CC$):            Dim C: For Each C In Split(CC): SetLcAsSum L, C:     Next: End Sub
Sub SetLccAsAvg(L As ListObject, CC$):            Dim C: For Each C In Split(CC): SetLcAsAvg L, C:     Next: End Sub
Sub SetLccLvl(L As ListObject, CC$, Lvl As Byte): Dim C: For Each C In SyzSS(CC): SetLcLvl L, C, Lvl: Next: End Sub
'---
Sub SetLcFml(L As ListObject, C, Fml$):           LcDta(L, C).Formula = Fml:                     End Sub
Sub SetLcAli(L As ListObject, C, A As eLofAli):  LcDta(L, C).HorizontalAliment = W1CvXlHAli(A): End Sub
Sub SetLcBdr(L As ListObject, C, A As eLofBdr):  BdrRg RgzLc(L, C), W1CvXlBdrIx(A):                       End Sub
Sub SetLcFmt(L As ListObject, C, Fmt$):           LcDta(L, C).NumberFormat = Fmt:                End Sub
Sub SetLcLvl(L As ListObject, C, Optional Lvl As Byte = 2):  EntLc(L, C).OutlineLevel = Lvl:                End Sub
Sub SetLcWdt(L As ListObject, C, W):              EntLc(L, C).ColumnWidth = W:                 End Sub
Sub SetLcWrp(L As ListObject, C, Wrp As Boolean): LcDta(L, C).WrapText = Wrp:                       End Sub
Sub SetLcTot(L As ListObject, C, A As eLofTot):  Lc(L, C).TotalsCalculation = W1CvXlTotCalc(A): End Sub
Sub SetLcCor(L As ListObject, C, Colr&):          LcDta(L, C).Interior.Color = Colr:                End Sub
Sub SetLcSum(L As ListObject, SumFld$, FmFld$, ToFld$): LcDta(L, SumFld).Formula = W1SumFldFml(FmFld, ToFld): End Sub
Sub SetLcLbl(L As ListObject, C, Lbl$):           W1LblCell(L, C).Value = Lbl:                           End Sub
Sub SetLcAsSum(L As ListObject, C): W1SetCalc L, C, xlTotalsCalculationSum:     End Sub
Sub SetLcAsCnt(L As ListObject, C): W1SetCalc L, C, xlTotalsCalculationCount:   End Sub
Sub SetLcAsAvg(L As ListObject, C): W1SetCalc L, C, xlTotalsCalculationAverage: End Sub
Private Sub W1SetCalc(L As ListObject, C, Calc As XlTotalsCalculation)
L.ShowTotals = True: Lc(L, C).TotalsCalculation = Calc
End Sub
Private Function W1LblCell(L As ListObject, F) As Range
Stop
End Function
Private Function W1SumFldFml$(FmFld$, ToFld$)
Stop
End Function
Private Function W1CvXlBdrIx(A As eLofBdr) As XlBordersIndex
Stop
End Function
Private Function W1CvXlHAli(A As eLofAli) As XlHAlign
Const CSub$ = CMod & "HAli"
Select Case A
Case "Left": W1CvXlHAli = xlHAlignLeft
Case "Right": W1CvXlHAli = xlHAlignRight
Case "Center": W1CvXlHAli = xlHAlignCenter
Case Else: Inf CSub, "Invalid Ali", "Valid Ali", LofAliss: Exit Function
End Select
End Function
Private Function W1CvXlTotCalc(A As eLofTot) As XlTotalsCalculation
Const CSub$ = CMod & "XTotCalc"
'Fm SACnt : "Sum | Avg | Cnt" @@
Dim O As XlTotalsCalculation
Select Case A
Case "Sum": O = xlTotalsCalculationSum
Case "Avg": O = xlTotalsCalculationAverage
Case "Cnt": O = xlTotalsCalculationCount
Case Else: Inf CSub, "Invalid TotCalcStr", "TotCalcStr Valid-TotCalcStr", A, "Sum Avg Cnt": Exit Function
End Select
W1CvXlTotCalc = O
End Function

'----
Sub SetLoFml(L As ListObject, Fmllny$())
Dim FmlLn: For Each FmlLn In Fmllny
    SetLcFmlln L, FmlLn
Next
End Sub
'-----------
Private Sub SetLcFmlln__Tst()
Dim Wb As Workbook: Set Wb = OpnFx(MB52Tp)
Dim Ws As Worksheet: Set Ws = Wb.Worksheets("MacauBchRat")
Dim L As ListObject: Set L = Ws.ListObjects(1)
SetLcFmlln L, "Litre=[@Btl] * [@Size] / 100"
SetLcFmlln L, "LitreHKD=[@Litre] * [@[MOP/Litre]] * [@[HKD/MOP]]"
SetLcFmlln L, "10%A=[@[XXX/Btl]] * [@Btl] * 0.1"
SetLcFmlln L, "10%B=[@Val] * 0.1"
SetLcFmlln L, "10%HKD=IF(ISBLANK([@[XXX/Btl]]),[@[10%B]],[@[10%A]])"
SetLcFmlln L, "HKD=[@LitreHKD] + [@[10%HKD]]"
SetLcFmlln L, "HKD/Ac=[@[Btl/Ac]]"
End Sub

Sub SetLcFmlln(L As ListObject, FmlLn)
With W2Brk(FmlLn)
    SetLcFml L, .S1, .S2
End With
End Sub

Private Function W2Brk(FmlLn) As S12
Dim A As S12: A = Brk1(FmlLn, "=")
If A.S1 = "" Or A.S2 = "" Then Thw CSub, "Invalid Fmlln", "Fmlln", FmlLn
With W2Brk
    .S1 = A.S1
    .S2 = "=" & A.S2
End With
End Function
