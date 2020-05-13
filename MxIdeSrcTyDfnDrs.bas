Attribute VB_Name = "MxIdeSrcTyDfnDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcTyDfnDrs."

Sub TyDfnDrsP__Tst()
BrwDrs TyDfnDrsP
End Sub

Function TyDfnDrsP() As Drs
':TyDfnDrs: :Drs-Mdn-Nm-Ty-Mem-Rmk
TyDfnDrsP = TyDfnDrszP(CPj)
End Function

Function TyDfnDrszP(P As VBProject) As Drs
TyDfnDrszP = DrszFF(TyDfnFF, TyDfnDyzP(P))
End Function

Function TyDfnDrszCmp(C As VBComponent) As Drs
Dim S$(): S = Src(C.CodeModule)
Dim Dy(): Dy = TyDfnDy(Vrmk(S), C.Name)
TyDfnDrszCmp = DrszFF(TyDfnFF, Dy)
End Function

Function TyDfnDyzP(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy TyDfnDyzP, TyDfnDyzM(C.CodeModule)
Next
End Function

Function TyDfnDyzM(M As CodeModule) As Variant()
TyDfnDyzM = TyDfnDy(Vrmk(Src(M)), Mdn(M))
End Function

Function TyDfnDy(VrmkBlk(), Mdn$) As Variant()
':TyDfnDy: :Dyo-Nm-Ty-Mem-Vrmk ! Fst-Ln must be :nn: :dd #mm# !rr
'                                ! Rst-Ln is !rr
'                                ! must term: nn dd mm, all of them has no spc
'                                ! opt      : mm rr
'                                ! :xx:     : should uniq in pj
Dim Vrmk: For Each Vrmk In Itr(VrmkBlk)
    PushI TyDfnDy, TyDfnDr(CvSy(Vrmk), Mdn)
Next
End Function

Private Function NotUse_TyDfnGp(Vrmk$()) As Variant()
Dim O(), Gp()
    Dim L: For Each L In Itr(Vrmk)
        Select Case True
        Case IsLnTyDfn(L)
            PushSomSi O, Gp         '<===
            Erase Gp
            PushI Gp, L             '<---
        Case IsLnTyDfnRmk(L)
            If Si(Gp) > 0 Then
                PushI Gp, L         '<--- Only with Fst-Ln, the Rst-Ln will be use, otherwise ign it.
            End If
        Case Else
            PushSomSi O, Gp         '<===
            Erase Gp                '<---
        End Select
    Next
NotUse_TyDfnGp = O
End Function

Private Sub TyDfnDr__Tst()
Dim Vrmk$()
GoSub YY
Exit Sub
YY:
    Vrmk = Sy("':Cell: :SCell-or-:WCell")
    Dmp TyDfnDr(Vrmk, "Md")
    Return
End Sub

Function TyDfnDr(TyDfnLy$(), Mdn$) As Variant()
'Assume: Fst Ln is ':nn: :dd [#mm#] [!rr]
'        Rst Ln is '                 !rr
Dim Dr(): Dr = TyDfnDrzLn1(TyDfnLy(0), Mdn)
Stop 'Dr(4) = AddNB(Dr(4), ExmRmkl(CvSy(RmvFstEle(TyDfnLy))))
TyDfnDr = Dr
End Function

Function TyDfnDrzLn1(FstTyDfnLin$, Mdn$) As Variant()
Const CSub$ = CMod & "TyDfnDrzLn1"
Dim L$: L = FstTyDfnLin
Dim Nm$, Ty$, Mem$, Rmk$
Nm = ShfTyDfnNm(L)
If Nm = "" Then Exit Function
Nm = RmvFstChr(Nm)
Ty = ShfColonTy(L)
Mem = ShfMemNm(L)
If L <> "" Then
    If FstChr(L) <> "!" Then Thw CSub, "Given FstTyDfnLin is not in valid format", "FstTyDfnLin", FstTyDfnLin
    Rmk = Trim(RmvFstChr(L))
End If
TyDfnDrzLn1 = Array(Mdn, Nm, Ty, Mem, Rmk)
End Function
