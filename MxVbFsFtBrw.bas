Attribute VB_Name = "MxVbFsFtBrw"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsFtBrw."

Sub VcFt(Ft)
ShellHid FmtQQ("Code.CMd ""?""", Ft)
End Sub

Sub NoteFt(Ft)
ShellMax FmtQQ("notepad.exe ""?""", Ft)
End Sub

Sub BrwFt(Ft, Optional UseVc As Boolean)
If UseVc Then
    VcFt Ft
Else
    NoteFt Ft
End If
End Sub
