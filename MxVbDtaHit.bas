Attribute VB_Name = "MxVbDtaHit"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str.Hit"
Const CMod$ = CLib & "MxVbDtaHit."
Function HitPfxAv(S, PfxAv()) As Boolean
If Si(PfxAv) = 0 Then HitPfxAv = 0: Exit Function
Dim I
For Each I In PfxAv
   If HasPfx(S, I) Then HitPfxAv = True: Exit Function
Next
End Function

Function HitPfxAp(S, ParamArray PfxAp()) As Boolean
Dim Av(): Av = PfxAp
HitPfxAp = HitPfxAv(S, Av)
End Function

Function HitPfxSpc(S, Pfx$, Optional B As VbCompareMethod = vbTextCompare) As Boolean
HitPfxSpc = HasPfx(S, Pfx & " ", B)
End Function

Function HitPfxySpc(S, Pfxy$(), Optional B As VbCompareMethod = vbTextCompare) As Boolean
Dim Pfx
HitPfxySpc = True
For Each Pfx In Pfxy
    If HasPfx(S, Pfx & " ", B) Then Exit Function
    If IsEqStr(S, Pfx, B) Then Exit Function
Next
HitPfxySpc = False
End Function

Function HasPfxzSomEle(Sy$(), Pfx$) As Boolean
Dim S
For Each S In Itr(Sy)
    If HasPfx(S, Pfx) Then
        HasPfxzSomEle = True
        Exit Function
    End If
Next
End Function
Function HitKss(S, Kss) As Boolean
HitKss = HitLikAy(S, SyzSS(Kss))
End Function

Function HitLikAy(S, LikAy$()) As Boolean
Dim Lik
For Each Lik In Itr(LikAy)
    If S Like Lik Then HitLikAy = True: Exit Function
Next
End Function

Function HitAp(V, ParamArray Ap()) As Boolean
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
HitAp = HasEle(Av, V)
End Function

Function HitNmStr(S, WhNmStr$) As Boolean
HitNmStr = HitNm(S, WhNm(WhNmStr))
End Function
Function HitNm(S, A As WhNm) As Boolean
HitNm = True
With A
If .IsEmp Then Exit Function
If HitLikAy(S, .ExlLikAy) Then HitNm = False: Exit Function
If HitRx(S, .Rx) Then Exit Function
If HitLikAy(S, .LikAy) Then Exit Function
End With
HitNm = False
End Function
Function HitAy(V, Ay) As Boolean
':HitXXX: :Fun-Patn ! #Hit-Somehting# HitXXX(V,XXX) As Boolean.
If Si(Ay) = 0 Then HitAy = True: Exit Function
HitAy = HasEle(Ay, V)
End Function

Private Sub HitPatn__Tst()
Dim A$, Patn$
Ept = True: A = "AA": Patn = "AA": GoSub Tst
Ept = True: A = "AA": Patn = "^AA$": GoSub Tst
Exit Sub
Tst:
    Act = HitPatn(A, Patn)
    C
    Return
End Sub

Function HitPatn(A, Patn) As Boolean
Static Re As New RegExp
Re.Pattern = Patn
HitPatn = Re.Test(A)
End Function
