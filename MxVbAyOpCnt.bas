Attribute VB_Name = "MxVbAyOpCnt"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Cnt"
Const CMod$ = CLib & "MxVbAyOpCnt."
Enum EmCntgOpt
    EiCntAll
    EiCntDup
    EiCntSng
End Enum

Function CntAy(Ay, Optional Opt As EmCntgOpt, Optional SrtOpt As EmCntSrtOpt, Optional By As eOrd) As String()
Const CSub$ = CMod & "CntAy"
Dim D As Dictionary: Set D = CntDi(Ay, Opt)
Dim K
Dim W%: W = DicItmWdt(D)
Dim O$()
Select Case SrtOpt
Case eNoSrt
    CntAy = CntLyzCntDi(D, W)
Case eSrtByCnt
    CntAy = QSrt(CntLyzCntDi(D, W), By)
Case eSrtByItm
    CntAy = CntLyzCntDi(SrtDic(D, By), W)
Case Else
    Thw CSub, "Invalid SrtOpt", "SrtOpt", SrtOpt
End Select
End Function

Sub BrwCnt(Ay, Optional Opt As EmCntgOpt)
Brw FmtCntDi(CntDi(Ay, Opt))
End Sub

Function CntgDrs(Ay, Opt As EmCntgOpt) As Drs
CntgDrs = DrszFF("Itm Cnt", CntgDy(Ay, Opt))
End Function

Function CntDyWhGt1zAy(Ay) As Variant()
CntDyWhGt1zAy = CntDyWhGt1(DyzDic(CntDi(Ay)))
End Function

Function CntgDy(Ay, Optional Opt As EmCntgOpt) As Variant()
CntgDy = DyzDic(CntDi(Ay, Opt))
End Function

Private Sub CntgDy__Tst()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = CntgDy(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
GoSub Tst
Exit Sub
Tst:
    Act = CntgDy(A)
    Ass IsEqAy(Act, Ept)
    Return
End Sub

Function SumSi&(Ay)
Dim I, O&
For Each I In Itr(Ay)
    O = O + Len(I)
Next
SumSi = O
End Function

Private Sub CntSiLin__Tst()
Debug.Print CntSiLin(SrczP(CPj))
End Sub

Function CntSiLin(Ay)
CntSiLin = "AyCntSi(" & Si(Ay) & "." & SumSi(Ay) & ")"
End Function
