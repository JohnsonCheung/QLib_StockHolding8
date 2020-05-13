Attribute VB_Name = "MxVbVarBrw"
Option Explicit
Option Compare Text
Const CNs$ = "Val.Brw"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbVarBrw."

Private Sub LisVal(V, Oup As OupOpt): LisAy Fmt(V), Oup: End Sub
Sub Vc(V, Optional FnPfx$): LisVal V, Vcg(FnPfx): End Sub
Sub B(V): Brw V: End Sub
Sub Brw(V, Optional FnPfx$): LisVal V, Brwg(FnPfx): End Sub

Function Fmt(V) As String()
Select Case True
Case IsStr(V):     Fmt = Sy(V)
Case IsLinesy(V): Fmt = FmtLinesy(CvSy(V))
Case IsArray(V):   Fmt = SyzAy(V)
Case IsAet(V):     Fmt = CvAet(V).Sy
Case IsDic(V):     Fmt = FmtDic(CvDic(V), InlValTy:=True)
Case IsEmpty(V):   Fmt = Sy("#Empty")
Case IsNothing(V): Fmt = Sy("#Nothing")
Case Else:         Fmt = Sy("#TypeName:" & TypeName(V))
End Select
End Function
