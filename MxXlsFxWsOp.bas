Attribute VB_Name = "MxXlsFxWsOp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsFxWsOp."

Sub ChkFxwEr(Fx$, W$, Er$())
If Si(Er) = 0 Then Exit Sub
Dim O$()
PushI O, "Excel File Path: [" & Pth(Fx) & "]"
PushI O, "Excel File     : [" & Fn(Fx) & "]"
PushI O, "Ws             : [" & W & "]"
BrwEr AddSy(O, Er)
End Sub
Private Sub BrwFxw__Tst()
BrwFxw MB52LasIFx
End Sub
Sub BrwFxw(Fx$, Optional Wsn0$)
BrwDrs DrszFxw(Fx, DftWsn(Wsn0, Fx))
End Sub
