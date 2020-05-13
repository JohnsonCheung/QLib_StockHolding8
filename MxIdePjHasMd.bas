Attribute VB_Name = "MxIdePjHasMd"
Option Explicit
Option Compare Text
Const CNs$ = "Pj.Prp"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdePjHasMd."
Function HasMd(P As VBProject, Mdn) As Boolean
HasMd = HasItn(P.VBComponents, Mdn)
End Function

Sub ChkMdnExist(P As VBProject, Mdn, Fun$)
If Not HasMd(P, Mdn) Then Thw Fun, "Should be a Mod", "Mdn MdTy", Mdn, ShtCmpTy(Cmp(Mdn).Type)
End Sub

Function HasMod(P As VBProject, Modn) As Boolean
If Not HasMd(P, Modn) Then Exit Function
HasMod = P.VBComponents(Modn).Type = vbext_ct_StdModule
End Function
