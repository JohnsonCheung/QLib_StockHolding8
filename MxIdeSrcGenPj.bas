Attribute VB_Name = "MxIdeSrcGenPj"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcGenPj."
':SrcPth:    :Pth ! #Src-Path#           is a :Pth.  Its fdr is a `{PjFn}.src`
':Distp:   :Pth ! #Distribution-Path#  is a :Pth.  It comes from :SrcPth by replacing .src to .dist
':Pthi: :Pth ! #Instance-Path#      of a @pth is any :TimNm :Fdr under @pth
':TimNm:   :Nm

Sub GenFxaFmCAcs() ' Gen a Fxa by CPj
GenFxaFmAcs Acs
End Sub

Sub GenFxaFmAcs(A As Access.Application)
ExpAcs A
GenFxa SrcPthzAcs(A)
End Sub

Sub GenFxa(SrcPth$)
Dim OFxa$:                   OFxa = DistFbai(SrcPth)
:                                   CrtFxa OFxa          ' <== Crt
Dim X As Excel.Application: Set X = NwXls
Dim P As VBProject:         Set P = XlsOpnXla(X, OFxa)
:                                   AddRfzSrcPth P, SrcPth ' <== Add Rf
:                                   LoadBas P, SrcPth    ' <== Load Bas
X.Quit
Inf CSub, "Fxa is created", "Fxa", OFxa
End Sub

Sub GenFba(SrcPth$)
Const CSub$ = CMod & "GenFba"
Dim OPj As VBProject
Dim OFba$:     OFba = DistFbai(SrcPth)     '#Oup-Fba#
:                     DltFfnIf OFba
:                     CrtFb OFba                    ' <== Crt OFba


:                     OpnFb Acs, OFba
            Set OPj = PjzAcs(Acs)
:                     AddRfzS OPj, RfSrczSrcPth(SrcPth) ' <== Add Rf
:                     LoadBas OPj, SrcPth             ' <== Load Bas
Dim Frm$():     Frm = FrmFfnAy(SrcPth)
Dim F: For Each F In Itr(Frm)
    Dim N$: N = RmvExt(RmvExt(F))
:               Acs.LoadFromText acForm, N, F       ' <== Load Frm
Next
#If False Then
'Following code is not able to save
Dim Vbe As Vbe: Set Vbe = Acs.Vbe
Dim C As VBComponent: For Each C In Acs.Vbe.ActiveVBProject.VBComponents
    C.Activate
    BoSavzV(Vbe).Execute
    Acs.Eval "DoEvents"
Next
#End If
MsgBox "Go access to save....."
Inf CSub, "Fba is created", "Fba", OFba
End Sub


Sub LoadBas(P As VBProject, Pth$)
Dim I: For Each I In Itr(BasFfnAy(Pth))
    P.VBComponents.Import I
Next
End Sub

Function BasFfnAy(Pth$) As String()
BasFfnAy = Ffny(Pth, "*.bas")
End Function

Sub GenFbaFmCAcs()
GenFbaFmAcs Acs
End Sub

Sub GenFbaFmAcs(A As Access.Application)
ExpAcs A                       ' <== Exp
GenFba SrcPthzP(PjzAcs(A))
End Sub
