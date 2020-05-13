Attribute VB_Name = "MxIdeMthIx"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthix."

Function Mthix&(Src$(), Mthn, Optional ShtMthTy$, Optional FmIx& = 0)
Dim I: For I = FmIx To UB(Src)
    If HitMth(Src(I), Mthn, ShtMthTy) Then
        Mthix = I: Exit Function
    End If
Next
Mthix = -1
End Function

Function HitMth(L, Mthn, ShtMthTy$) As Boolean
Dim A As Mthn3: A = Mthn3zL(L)
If Mthn <> A.Nm Then Exit Function
If HitOptEq(A.ShtMdy, ShtMthTy) Then HitMth = True: Exit Function
Debug.Print FmtQQ("HitMth: Mthn[?] Hits L but mis match given ShtMthTy[?].  Act ShtMthTy[?].  Ln=[?]", A.Nm, A.ShtTy, ShtMthTy, L)
End Function

Function FstMthix&(Src$(), Optional Fm = 0)
For FstMthix = Fm To UB(Src)
    If IsMthln(Src(FstMthix)) Then Exit Function
Next
FstMthix = -1
End Function

Function FstMthlno&(Md As CodeModule)
Dim J&: For J = 1 To Md.CountOfLines
   If IsMthln(Md.Lines(J, 1)) Then
       FstMthlno = J
       Exit Function
   End If
Next
End Function


Function MthLcntByIx(Src$(), Mthix&) As Lcnt
MthLcntByIx = LcntByBei(Mthix, SrcEix(Src, Mthix))
End Function

Function MthixItr(Src$())
Asg Itr(Mthixy(Src)), MthixItr
End Function

Private Sub Mthixy__Tst()
Dim S$()
GoSub Z
Exit Sub
Z:
    S = SrczMdn("MxMthix")
    Dim MIxy&(): MIxy = Mthixy(S)
    Brw AwIxy(S, MIxy)
    Return

End Sub

Function Mthixy(Src$()) As Long() ' method index array
Dim Ix&: For Ix = 0 To UB(Src)
    If IsMthln(Src(Ix)) Then
        PushI Mthixy, Ix
    End If
Next
End Function

Sub Mthlno__Tst()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MthnyzM(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, Mthlno(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Si(Ny), "Z_Mthlno"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
BrwAy O
End Sub

Function MthlnyzMN(M As CodeModule, Mthn) As Long()
MthlnyzMN = AmInc(MthixyzSNT(Src(M), Mthn))
End Function

Function MthixyzSNT(Src$(), Mthn, Optional ShtMthTy$) As Long()
Dim Ix&: Ix = Mthix(Src, Mthn, ShtMthTy): If Ix = -1 Then Exit Function
PushI MthixyzSNT, Ix
If IsPrpln(Src(Ix)) Then
    PushIx MthixyzSNT, Mthix(Src, Mthn, ShtMthTy, Ix + 1)
End If
End Function

Function MthixyzN(Src$(), Mthny$()) As Long()
Dim Ix: For Each Ix In MthixItr(Src)
    Dim L$: L = Src(Ix)
    Dim N$: N = Mthn(L)
    If HasEle(Mthny, N) Then PushI MthixyzN, Ix
Next
End Function

Function MthixzMN&(M As CodeModule, Mthn, Optional ShtMthTy$)
MthixzMN = MthixzN(Src(M), Mthn, ShtMthTy)
End Function

Function MthixzN&(Src$(), Mthn, Optional ShtMthTy$)
Dim Ix&
For Ix = 0 To UB(Src)
    With Mthn3zL(Src(Ix))
        If .Nm = Mthn Then
            If HitOptEq(.ShtTy, ShtMthTy) Then
                MthixzN = Ix
                Exit Function
            End If
            Debug.Print FmtQQ("MthixzN: Given Mthn[?] Hit, not given ShtMthTy[?].  Act ShtMth[?]", Mthn, ShtMthTy, .ShtTy)
            If .ShtTy = "???" Then Stop
        End If
    End With
Next
MthixzN = -1
End Function

Function HitOptEq(S, OptEq$) As Boolean ' If OptEq="" always return True, else return S=OptEq
If OptEq = "" Then HitOptEq = True: Exit Function
HitOptEq = S = OptEq
End Function

Function MthlnozCLno&(M As CodeModule, CLno&)
Dim L&: For L = CLno To 1 Step -1
    If IsMthln(M.Lines(L, 1)) Then MthlnozCLno = L: Exit Function
Next
End Function

Function Mthlno&(M As CodeModule, MthNm, Optional FmLno& = 1)
'Mthlno:: :Lno Method line number
'Lno::    :No  Line number
'No::     :Lng a number is running from 1
Dim O&: For O = FmLno To M.CountOfLines
    If Mthn(M.Lines(O, 1)) = MthNm Then Mthlno = O: Exit Function
Next
End Function

Function MthLcnt2(M As CodeModule, Mthn) As Lcnt2
MthLcnt2.A = MthLcnt(M, Mthn)
Stop
End Function

Function MthLcnt(M As CodeModule, Mthn, Optional FmLno& = 1) As Lcnt
Dim Lno&: Lno = Mthlno(M, Mthn, FmLno)
If Lno = 0 Then Thw CSub, "@Mthn not found in @Md", "Mdn Mthn", Mdn(M), Mthn
With MthLcnt
    .Lno = Lno
    Dim E&: E = SrcEno(M, 1):    If E = 0 Then Thw CSub, "@Mthn has a Mthlno but no SrcEno", "@Mthn Mthlno", Mthn, Lno
    .Cnt = E - Lno + 1
End With
End Function
