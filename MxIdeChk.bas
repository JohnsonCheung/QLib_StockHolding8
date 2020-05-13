Attribute VB_Name = "MxIdeChk"
Option Explicit
Option Compare Text
Const CNs$ = "Chk"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeChk."
Function ChkMdn(P As VBProject, Mdn) As Boolean
If HasMd(P, Mdn) Then Exit Function
MsgBox FmtQQ("Mdn not found: ?|In Pj: ?", Mdn, P.Name), vbCritical
ChkMdn = True
End Function
