Attribute VB_Name = "MxXlsLoFmtLof"
Option Explicit
Option Compare Text
Const CNs$ = "Lof"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLoFmtLof."

':Lof: :Ly #ListObject-Formatter# ! Each line is Ly with T1 LofLofT1nn"
':FldLikss: :Likss #Fld-Lik-SS# ! A :SS to expand a given Fny
':Ali:   :
Sub FmtLoByLofLy(Lo As ListObject, LofLy$())

End Sub
Sub FmtLo(Lo As ListObject, A As Lof)
Dim J%
With A
    SetLon Lo, .Lon
    For J = 0 To LofAliUB(.Ali): FmtAli Lo, .Ali(J): Next
    For J = 0 To LofBdrUB(.Bdr): FmtBdr Lo, .Bdr(J): Next
    For J = 0 To LofCorUB(.Cor): FmtCor Lo, .Cor(J): Next
    For J = 0 To LofFmlUB(.Fml): FmtFml Lo, .Fml(J): Next
    For J = 0 To LofFmtUB(.Fmt): FmtFmt Lo, .Fmt(J): Next
    For J = 0 To LofLblUB(.Lbl): FmtLbl Lo, .Lbl(J): Next
    For J = 0 To LofLvlUB(.Lvl): FmtLvl Lo, .Lvl(J): Next
    For J = 0 To LofSumUB(.Sum): FmtSum Lo, .Sum(J): Next
    For J = 0 To LofTotUB(.Tot): FmtTot Lo, .Tot(J): Next
    For J = 0 To LofWdtUB(.Wdt): FmtWdt Lo, .Wdt(J): Next
    FmtTit Lo, .Tit
End With
End Sub

Sub AddLoFml(L As ListObject, ColNm$, Fml$)
Dim O As ListColumn
Set O = L.ListColumns.Add
O.Name = ColNm
O.DataBodyRange.Formula = Fml
End Sub



Sub FmtTotLnk(L As ListObject, C)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = L.ListColumns(C).DataBodyRange
Set Ws = WszRg(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, NRoZZRg(R) + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Sub FmtTit(L As ListObject, A() As LofTit)
Dim Sq(), R As Range
    Sq = TitSq(A): If Si(Sq) = 0 Then Exit Sub
    Set R = XTitAt(L, UBound(Sq(), 1))
Set R = RgzSq(Sq(), R)
XMgeTit R
BdrInside R
BdrAround R
End Sub

Sub XMgeTit(TitRg As Range)
Dim J%
For J = 1 To NRoZZRg(TitRg)
    XMgeTitH RgR(TitRg, J)
Next
For J = 1 To TitRg.Columns.Count
    XMgeTitV RgC(TitRg, J)
Next
End Sub

Sub XMgeTitH(TitRg As Range)
TitRg.Application.DisplayAlerts = False
Dim J%, C1%, C2%, V, LasV
LasV = RgRC(TitRg, 1, 1).Value
C1 = 1
For J = 2 To TitRg.Columns.Count
    V = RgRC(TitRg, 1, J).Value
    If V <> LasV Then
        C2 = J - 1
        If Not IsEmpty(LasV) Then
            RgRCC(TitRg, 1, C1, C2).MergeCells = True
        End If
        C1 = J
        LasV = V
    End If
Next
TitRg.Application.DisplayAlerts = True
End Sub

Sub XMgeTitV(A As Range)
Dim J%
For J = NRoZZRg(A) To 2 Step -1
    MgeCellAbove RgRC(A, J, 1)
Next
End Sub

Function XTitAt(Lo As ListObject, NTitRow%) As Range
Set XTitAt = RgRC(Lo.DataBodyRange, 0 - NTitRow, 1)
End Function

Function TitSq(A() As LofTit) As Variant()
Dim Fny$()
Dim Col()
    Dim F$, I, Tit$
    For Each I In Fny
        F = I
'        Tit = FstElezRmvT1(TitLy, F)
        If Tit = "" Then
            PushI Col, Sy(F)
        Else
            PushI Col, AmTrim(SplitVBar(Tit))
        End If
    Next
TitSq = Transpose(SqzDy(Col))
End Function

Private Sub TitSq__Tst()
Dim Tit() As LofTit, Fny$()
'----
Dim A$(), Act(), Ept()
'TitLy
    Erase A
    Push A, "A A1 | A2 11 "
    Push A, "B B1 | B2 | B3"
    Push A, "C C1"
    Push A, "E E1"
    Tit = LofTitAy(A)

Fny = SyzSS("A B C D E")
Ept = TitSq(Tit)
    SetSqr Ept, 1, SyzSS("A1 B1 C1 D E1")
    SetSqr Ept, 2, Array("A2 11", "B2")
    SetSqr Ept, 3, Array(Empty, "B3")
GoSub Tst
Exit Sub
'---
'Tit
    Erase A
    PushI A, "A AAA | skldf jf"
    PushI A, "B skldf|sdkfl|lskdf|slkdfj"
    PushI A, "C askdfj|sldkf"
    PushI A, "D fskldf"
    Tit = LofTitAy(A)
BrwSq TitSq(Tit)

Exit Sub
Tst:
    Act = TitSq(Tit)
    Ass IsEqSq(Act, Ept)
    Return
End Sub

Sub FmtAli(L As ListObject, A As LofAli): Dim F: For Each F In Itr(A.Fny): SetLcAli L, F, A.Ali: Next: End Sub

Function FnyzT1FldLikss(Fny$(), LinOf_T1_FldLikss) As String()
FnyzT1FldLikss = AwLikss(Fny, RmvT1(LinOf_T1_FldLikss))
End Function


Sub FmtBet(L As ListObject, LinOf_Sum_Fm_To)
Dim FSum$, FFm$, FTo$: AsgTTRst LinOf_Sum_Fm_To, FSum, FFm, FTo
EntLc(L, FSum).Formula = FmtQQ("=Sum([?]:[?])", FFm, FTo)
End Sub

Private Sub FmtFml(L As ListObject, A As LofFml): EntLc(L, A.Fld).Formula = A.Fml:                       End Sub
Private Sub FmtLbl(L As ListObject, A As LofLbl): SetLcLbl L, A.Fld, A.Lbl:                                    End Sub
Private Sub FmtSum(L As ListObject, A As LofSum): SetLcSum L, A.SumFld, A.FmFld, A.ToFld:                      End Sub
Private Sub FmtBdr(L As ListObject, A As LofBdr): Dim F: For Each F In Itr(A.Fny): SetLcBdr L, F, A.Bdr: Next: End Sub
Private Sub FmtLvl(L As ListObject, A As LofLvl): Dim F: For Each F In Itr(A.Fny): SetLcLvl L, F, A.Lvl: Next: End Sub
Private Sub FmtCor(L As ListObject, A As LofCor): Dim F: For Each F In Itr(A.Fny): SetLcCor L, F, A.Cor: Next: End Sub
Private Sub FmtFmt(L As ListObject, A As LofFmt): Dim F: For Each F In Itr(A.Fny): SetLcFmt L, F, A.Fmt: Next: End Sub
Private Sub FmtTot(L As ListObject, A As LofTot): Dim F: For Each F In Itr(A.Fny): SetLcTot L, F, A.Tot: Next: End Sub
Private Sub FmtWdt(L As ListObject, A As LofWdt): Dim F: For Each F In Itr(A.Fny): SetLcWdt L, F, A.Wdt: Next: End Sub



'Tst-------------------------------------------------------------
Private Sub FmtLo__Tst()
Dim Lo As ListObject, Fmtr() As String 'Lofr
'------------
Set Lo = SampLo
Fmtr = SampLofLy
GoSub Tst
Exit Sub
Tst:
    FmtLoByLofLy Lo, Fmtr
    Return
End Sub

Private Sub FmtBdr__Tst()
Dim Ln$, Lo As ListObject, A As LofBdr
'--
Set Lo = SampLo
'--
GoSub T1
GoSub T2
Exit Sub
T1: Ln = "Left A B C": GoTo Tst
T2: Ln = "Left D E F": GoTo Tst
T3: Ln = "Right A B C": GoTo Tst
T4: Ln = "Center A B C": GoTo Tst
Tst:
    FmtBdr Lo, A      '<=='
    Stop
    Return
End Sub


'Fun===========================================================================
Function LoHdrCell(L As ListObject, C) As Range
Set LoHdrCell = A1zRg(CellAbove(L.ListColumns(C).Range))
End Function

Sub StdFmtWbLo(B As Workbook)
Dim S As Worksheet
For Each S In B.Sheets
    StdFmtLozWs S
Next
End Sub
Sub StdFmtLozWs(S As Worksheet)
Dim L As ListObject: For Each L In S.ListObjects
    FmtLo L, StdLof
Next
End Sub

Sub FmtLoByStdLof(L As ListObject)
FmtLo L, StdLof
End Sub


Function StdLof() As Lof

End Function
