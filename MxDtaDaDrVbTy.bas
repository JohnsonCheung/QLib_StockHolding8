Attribute VB_Name = "MxDtaDaDrVbTy"
Option Explicit
Option Compare Text
Const CNs$ = "Dta.VbTy"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaDrVbTy."

Function VbTyTslzDr$(Dr)
Dim O$()
Dim V: For Each V In Itr(Dr)
    PushI O, TypeName(V)
Next
VbTyTslzDr = JnTab(O)
End Function

Function FstDr(Dy())
If Si(Dy) = 0 Then Exit Function
Dim Dr: Dr = Dy(0)
Dim J%: For J = Si(Dr) To NColzDy(Dy) - 1
    PushI Dr, ""
Next
FstDr = Dr
End Function
