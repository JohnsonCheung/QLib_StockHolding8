Attribute VB_Name = "MxIdeSrcCnstDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CNs$ = "Cnst"
Const CMod$ = CLib & "MxIdeSrcCnstDrs."
Public Const CnstFF$ = "Mdn Mdy Cnstn TyChr AftEq"

Function CnstDrsP() As Drs
CnstDrsP = CnstDrszP(CPj)
End Function

Function CnstDrszP(P As VBProject) As Drs
Dim O As Drs
Dim C As VBComponent: For Each C In P.VBComponents
    O = AddDrs(O, CnstDrszM(C.CodeModule))
Next
CnstDrszP = O
End Function

Function CnstDrs(Dcl$(), Mdn$) As Drs
CnstDrs = DrszFF(CnstFF, CnstDy(Dcl, Mdn))
End Function

Function CnstDy(Dcl$(), Mdn$) As Variant()
Dim L: For Each L In Itr(Dcl)
    PushSomSi CnstDy, CnstDr(L, Mdn)
Next
End Function

Function CnstDr(Ln, Optional Mdn$) As Variant()
'Ret    : :Dro|EmpAv if @Ln is not a cnst-cont-Ln
Dim L$: L = Ln
Dim Mdy$: Mdy = ShfMdy(L)               '<-- 1 Mdy
    Select Case Mdy
    Case "Public": Mdy = "Pub"
    Case "", "Private": Mdy = ""
    Case Else: Exit Function            '<===
    End Select

                    If Not ShfCnst(L) Then Exit Function
Dim Cnstn$: Cnstn = ShfNm(L)                '<-- 2 Nm
                    If Cnstn = "" Then Exit Function '<==
Dim TyChr$: TyChr = ShfTyChr(L)             '<-- 3 TyChr
                    If Not ShfPfx(L, " = ") Then Exit Function  '<==
          CnstDr = Array(Mdn, Mdy, Cnstn, TyChr, L)
End Function

Function CnstDrszM(M As CodeModule) As Drs
CnstDrszM = CnstDrs(Dcl(M), Mdn(M))
End Function
