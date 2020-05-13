Attribute VB_Name = "gzFmt15MthTit"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzFmt15MthTit."
Private Sub FmtFcDteTit__Tst()
FmtFcDteTit SampShWb, YM(18, 7)
MaxiWb SampShWb
End Sub
Private Sub FmtSdDteTit__Tst()
FmtSdDteTit SampShWb, YM(19, 12)
MaxiWb SampShWb
End Sub
'== Fc=Forecast
'== Sd=StkDays
Function FmtFcDteTit(Wb As Workbook, A As YM, Optional NMth% = 15)
Dim Sq(): Sq = FcDteTitSq(A)
Dim Rg As Range

'Without doing this Rg.Merge and Rg.UnMerge will break

MinvWb Wb
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If IsFc(Ws.Name) Then
        Set Rg = FcDteTitRg(Ws, NMth)
        Rg.UnMerge
        Rg.Value = Sq
        MgeTit Rg
    End If
Next
End Function
Function FmtSdDteTit(Wb As Workbook, A As YM)
Dim Sq(): Sq = SdDteTitSq(A)
Dim Rg As Range

'Without doing this Rg.Merge and Rg.UnMerge will break
MinvWb Wb
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If IsSd(Ws.Name) Then
        Set Rg = SdDteTitRg(Ws)
        Rg.UnMerge
        Rg.Value = Sq
        MgeTit Rg
    End If
Next
End Function
'---=========================================================================================== A_MgeTit
Sub UnMgeRg(Rg As Range)
MinvRg Rg
Rg.UnMerge
End Sub
Sub MgeRg(Rg As Range)
MinvRg Rg
Rg.Merge
End Sub
Sub MgeTit(Tit As Range)
MinvRg Tit
WszRg(Tit).Activate
Tit.UnMerge
Dim R%: For R = 1 To Tit.Rows.Count
    MgeRow RgR(Tit, R)
Next
Tit.BorderAround XlLineStyle.xlContinuous, xlThick
With Tit.Borders(xlInsideHorizontal)
    .LineStyle = XlLineStyle.xlContinuous
    .Weight = XlBorderWeight.xlThick
End With
With Tit.Borders(xlInsideVertical)
    .LineStyle = XlLineStyle.xlContinuous
    .Weight = XlBorderWeight.xlThick
End With
End Sub
Private Sub MgeRow(Row As Range)
Dim I: For Each I In Itr(MgableAy(Row))
    'Tit.Application.WindowState = xlMaximized  'Need this, other Merge will break
    CvRg(I).Merge
    'Tit.Application.WindowState = xlMinimized
    CvRg(I).HorizontalAlignment = XlHAlign.xlHAlignCenter
Next
End Sub
Private Function MgableAy(Row As Range) As Range()
Dim Dr(): Dr = FstDrzRg(Row)
Dim Ay() As C12: Ay = MgableC12Ay(Dr)
Dim RowC%: RowC = Row.Column
Dim Ws As Worksheet: Set Ws = Row.Parent
Dim R&: R = Row.Row
Dim J%: For J = 0 To C12UB(Ay)
    Dim M As C12: M = Ay(J)
    Dim C1%: C1 = M.C1 + RowC
    Dim C2%: C2 = M.C2 + RowC
    PushObj MgableAy, WsRCC(Ws, R, C1, C2)
Next
End Function
Sub ChkRow(Rg As Range)
If Rg.Rows.Count <> 1 Then Stop
End Sub

Private Function ToMgeRg(MgeTitRow As Range) As Range
'Return the fst ToMgeRg withing MgeTitRow
Dim Ws As Worksheet: Ws = WszRg(MgeTitRow)
Dim Row As RCC: Row = RCCzRg(MgeTitRow)
Dim DtaCno1%: DtaCno1 = FstDtaCno(Ws, Row)
If DtaCno1 = 0 Then Exit Function
Dim DtaCno2%: DtaCno2 = FstDtaCno(Ws, NxtRCC(Row))
If DtaCno2 = 0 Then Exit Function
Set ToMgeRg = RgRCC(MgeTitRow, 1, DtaCno1, DtaCno2 - 1)
End Function

Private Function MgableC12Ay(Dr()) As C12()
Dim C%: For C = 0 To UB(Dr) - 1 ' Mgable C must have next, so -1
    If IsMgable(Dr, C) Then
        PushC12 MgableC12Ay, C12(C, MgableC2(Dr, C))
    End If
Next
End Function
Private Function IsMgable(Dr(), C) As Boolean
If IsEmpty(Dr(C)) Then Exit Function
IsMgable = IsEmpty(Dr(C + 1))
End Function

Private Function MgableC2%(Dr(), C)
Dim U&: U = UB(Dr)
Dim O%: For O = C + 1 To U
    If Not IsEmpty(Dr(O)) Then MgableC2 = O - 1: Exit Function
Next
MgableC2 = U
End Function

'---============================================================================================== A_DteTitRg
Private Sub Z1()
MsgBox SdDteTitRg(SampSdWs).Address
MsgBox FcDteTitRg(SampFcWs).Address
End Sub
Private Function SdDteTitRg(SdWs As Worksheet, Optional NMth% = 15) As Range: Set SdDteTitRg = DteTitRg(FstLo(SdWs), "StkDays01", NMth, 2, 1): End Function
Private Function FcDteTitRg(FcWs As Worksheet, Optional NMth% = 15) As Range: Set FcDteTitRg = DteTitRg(FstLo(FcWs), "M01", NMth): End Function
Private Function DteTitRg(Lo As ListObject, Dte01ColNm$, NMth%, Optional NColPerMth% = 1, Optional NSpcRow%) As Range
':DteTitRg: ! #Dte-title-range# is a 2-Rows-N-Months-range above DteTitAt having @NSpcRow
Dim R1%, R2%, C1%, C2%
R1 = -1 - NSpcRow
R2 = 0 - NSpcRow
C1 = 1
C2 = NMth * NColPerMth
Dim Dte01Cell As Range: Set Dte01Cell = LcHdrCell(Lo, Dte01ColNm)
Set DteTitRg = RgRCRC(Dte01Cell, R1, C1, R2, C2)
End Function

'---============================================================================================== DteSq
Private Function FcDteTitSq(A As YM) As Variant(): FcDteTitSq = DteTitSq(A, 1): End Function
Private Function SdDteTitSq(A As YM) As Variant(): SdDteTitSq = DteTitSq(A, -1, True): End Function
Private Sub DteTitSq__Tst()
Dim S As YM: S = YM(17, 4)
Dim N(): N = DteTitSq(S)           'Normal
Dim P(): P = DteTitSq(S, -1)       'Previous
Dim N2(): N2 = DteTitSq(S, , True) 'Normal Double
Dim P2(): P2 = DteTitSq(S, -1, True)   'Prevous Double
Stop
End Sub
Private Function DteTitSq(A As YM, Optional Direction% = 1, Optional IsDouble As Boolean, Optional NMth% = 15) As Variant()
'@IsDouble: is the date title double column?
'@Direction: -1 means back & 1 means forecast
Dim MAy() As Date: MAy = MthAy(A.Y, A.M, Direction, NMth)
Dim NC%: NC = NMth: If IsDouble Then NC = NC * 2
Dim O(): ReDim O(1 To 2, 1 To NC)
Dim C%, D As Date, M$, Y%
If IsDouble Then
    For C = 1 To NMth
        D = MAy(C - 1)
        Y = Year(D)
        M = Format(D, "MMM")
        O(1, C * 2 - 1) = Y
        O(1, C * 2) = Y
        O(2, C * 2 - 1) = M
        O(2, C * 2) = M
    Next
Else
    For C = 1 To NMth
        D = MAy(C - 1)
        Y = Year(D)
        M = Format(D, "MMM")
        O(1, C) = Y
        O(2, C) = M
    Next
End If
DteTitSq = RmvIfSamAsPrvCol(O)
End Function

Private Function RmvIfSamAsPrvCol(Sq()) As Variant()
Dim O(): O = Sq
Dim R%: For R = 1 To UBound(Sq, 1)
    Dim C%: For C = UBound(Sq, 2) To 2 Step -1
        If O(R, C) = O(R, C - 1) Then O(R, C) = Empty
    Next
Next
RmvIfSamAsPrvCol = O
End Function
