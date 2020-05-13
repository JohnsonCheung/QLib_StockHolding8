Attribute VB_Name = "MxVbRunObjMth"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbRunObjMth."
Const CNs$ = "Run"
Sub RunItoMth(Ito, ObjMth)
Dim Obj As Object: For Each Obj In Ito
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Sub RunOyMth(Oy, ObjMth)
Dim Obj: For Each Obj In Itr(Oy)
    CallByName Obj, ObjMth, VbMethod
Next
End Sub
