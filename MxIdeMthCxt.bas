Attribute VB_Name = "MxIdeMthCxt"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthCxt."
Function MthCxtLcntAy(Src$(), Mthn) As Lcnt()

End Function

Function MthCxtLcntAyFmMd(M As CodeModule, Mthn) As Lcnt()
End Function

Function MthCxtLcntAyByMth(Src$(), Mth() As Lcnt) As Lcnt()
Dim J%: For J = 0 To LcntUB(Mth)
    PushLcnt MthCxtLcntAyByMth, MthCxtLcnt(Src, Mth(J))
Next
End Function

Function MthCxtLcnt(Src$(), Mth As Lcnt) As Lcnt
With Mth
Dim N%: N = NContln(Src, .Lno)
MthCxtLcnt = Lcnt(.Lno - N, .Cnt - N - 1)
End With
End Function

Function IsRmkdMdLcnt(M As CodeModule, A As Lcnt) As Boolean
IsRmkdMdLcnt = IsRmkdSrc(SrcByLcnt(M, A))
End Function

Function IsRmkdMthLy(Mthly$()) As Boolean
IsRmkdMthLy = IsRmkdSrc(MthCxt(Mthly))
End Function

Function MthCxt(Mthly$()) As String()
Const CSub$ = CMod & "MthCxt"
ChkIsMthLy Mthly, CSub
Dim J%: For J = NxtSrcIx(Mthly) To UB(Mthly) - 1
    PushI MthCxt, Mthly(J)
Next
End Function

Sub ChkIsMthLy(Ly$(), Fun$)
Select Case True
Case Si(Ly) <= 1
Case Not IsMthln(Ly(0)): Thw CSub, "Fst-Ln of Ly is not Mth", "@Ly", Ly
End Select
End Sub
