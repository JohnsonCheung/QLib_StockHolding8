Attribute VB_Name = "MxVbFsFno"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsFno."
Function FnoRnd128%(Ffn)
FnoRnd128 = FnoRnd(Ffn, 128)
End Function

Function FnoRnd%(Ffn, RecLen&)
Dim O%: O = FreeFile(1)
Open Ffn For Random As #O Len = RecLen
FnoRnd = O
End Function

Function FnoBin%(Ffn)
Dim O%: O = FreeFile(1)
Open Ffn For Binary As #O
FnoBin = O
End Function

Function FnoA%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
FnoA = O
End Function

Function FnoI%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FnoI = O
End Function

Function FnoO%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FnoO = O
End Function
