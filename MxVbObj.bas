Attribute VB_Name = "MxVbObj"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Obj"
Const CMod$ = CLib & "MxVbObj."
Const DoczP$ = "Prpp."
Const DoczPn$ = "PrpNm."
Enum EmThw
    EiThwEr
    EiNoThw
End Enum
Function IsEqObj(A, B) As Boolean
IsEqObj = ObjPtr(A) = ObjPtr(B)
End Function

Function IsEqVar(A, B) As Boolean
IsEqVar = VarPtr(A) = VarPtr(B)
End Function

Function IntozOy(Into, Oy)
Erase Into
Dim O, I
For Each I In Itr(Oy)
    PushObj Into, I
Next
End Function

Function LngAyzOyPrp(Oy, Prpp$) As Long()
LngAyzOyPrp = CvLngAy(IntozOyPrp(EmpLngAy, Oy, Prpp))
End Function

Function IntozOyPrp(Into, Oy, Prpp$)
Dim O: O = NwAy(Into)
Dim Obj: For Each Obj In Itr(Oy)
    Push O, Opv(Obj, Prpp)
Next
IntozOyPrp = O
End Function

Function AddObjOy(Obj As Object, Oy)
Dim O: O = Oy
Stop
Erase O
PushObj O, Obj
PushObjAy O, Oy
AddObjOy = O
End Function

Sub ChkNothing(A, Fun$)
If IsNothing(A) Then Thw Fun, "Given object is nothing"
End Sub
Function Objn$(A)
ChkNothing A, CSub
Objn = A.Name
End Function
