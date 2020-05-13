Attribute VB_Name = "MxIdeMthOpRmv"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Op"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthOpRmv."
Sub RplMth(M As CodeModule, Mthn, NewL$)
If HasCdl(M, NewL) Then Exit Sub
DltMth M, Mthn
M.AddFromString NewL
End Sub

Function DltMthny(M As CodeModule, Mthny$()) As Boolean
Dim O As Boolean
Dim Mthn: For Each Mthn In Itr(Mthny)
    If DltMth(M, Mthn) Then O = True
Next
DltMthny = O
End Function

Sub DltMth__Tst()
DltMth CMd, "AA"
End Sub
Function DltMth(M As CodeModule, Mthn) As Boolean ' Dlt fst mth in Md if exist else inf mth not fnd
Dim Bei As Bei: Bei = MthBei(Src(M), Mthn)
If IsEmpBei(Bei) Then Debug.Print "DltMth: Given Mthn not found in Md[" & M.Parent.Name & "] Mthn[" & Mthn & "]": Exit Function
DltCdlzLcnt M, LcntzBei(Bei)
DltMth = True
End Function
Private Sub DltCdlzLcnt(M As CodeModule, A As Lcnt)
DltCdl M, A.Lno, A.Cnt
End Sub
Private Function IsEmpBei(A As Bei) As Boolean
Select Case True
Case A.Bix < 0, A.Eix < 0: IsEmpBei = True
End Select
End Function
Private Function MthBei(Src$(), Mthn, Optional ShtMthTy$) As Bei
Dim B&: B = Mthix(Src, Mthn, ShtMthTy)
MthBei = Bei(B, MthEix(Src, B))
End Function
Private Function Mthix&(Src$(), Mthn, Optional ShtMthTy$, Optional FmIx& = 0)
Dim I: For I = FmIx To UB(Src)
    If HitMth(Src(I), Mthn, ShtMthTy) Then
        Mthix = I: Exit Function
    End If
Next
Mthix = -1
End Function
Private Function HitMth(L, Mthn, ShtMthTy$) As Boolean

Dim A As Mthn3: A = Mthn3zL(L)
If Mthn <> A.Nm Then Exit Function
If HitOptEq(A.ShtMdy, ShtMthTy) Then HitMth = True: Exit Function
Debug.Print "HitMth: Mthn[" & Mthn & "] Hits L but mis match given ShtMthTy[" & ShtMthTy & "].  Act ShtMthTy[" & ShtMthTy & "].  Ln=[" & L & "]"

End Function

Sub EnsMth__Tst()
Dim Mthl$: Mthl = "Private Sub AA(): End Sub" & vbCrLf & "Private Sub BB(): End Sub"
EnsMth CMd, Mthl, Sy("AA BB")
End Sub

Sub EnsMth(M As CodeModule, MthCdl$, Mthny$()) '@M should have this @MthCdl, otherwise, dlt @Mthny and ins @MthCdl
If HasCdl(M, MthCdl) Then Exit Sub
Dim P As VBProject: Set P = PjzM(M)
Dim Ft1$, Ft2$, H$, N$
    N = Mdn(M)
    H = TmpFdr("EnsMth")
    Ft1 = H & N & ".DltMthny.txt"
    Ft2 = H & N & ".Append.MthCdl.txt"
WrtAy Mthny, Ft1
WrtStr MthCdl, Ft2
End Sub

Sub EnsMthPass2()

End Sub

