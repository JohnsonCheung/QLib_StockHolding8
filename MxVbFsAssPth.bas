Attribute VB_Name = "MxVbFsAssPth"
Option Explicit
Option Compare Text
Const CNs$ = "AssPth"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbFsAssPth."

Function EnsAssPth$(Ffn)
Dim O$: O = AssPth(Ffn)
EnsPth O
EnsAssPth = O
End Function

Function AssPth$(Ffn)
AssPth = Pth(Ffn) & "." & Fn(Ffn) & "\"
End Function
