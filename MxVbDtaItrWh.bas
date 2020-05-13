Attribute VB_Name = "MxVbDtaItrWh"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbDtaItrWh."

Function IntozIwEq(Into, Itr, Prpp, V)
IntozIwEq = NwAy(Into)
Dim Obj: For Each Obj In Itr
    If Opv(Obj, Prpp) = V Then PushObj IntozIwEq, Obj
Next
End Function

Function IwNm(Itr, B As WhNm)
IwNm = ItrClnAy(Itr)
Dim O
For Each O In Itr
    If HitNm(Objn(O), B) Then
        Push IwNm, O
    End If
Next
End Function

Function IwPrpTrue(Itr, TruePrpp)
IwPrpTrue = ItrClnAy(Itr)
Dim Obj: For Each Obj In Itr
    If Opv(Obj, TruePrpp) Then
        Push IwPrpTrue, Obj
    End If
Next
End Function
