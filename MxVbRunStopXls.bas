Attribute VB_Name = "MxVbRunStopXls"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxVbRunStopXls."

Sub StopXls()
EnsFt StopXlsFps, StopXlsPsCxt
Shell StopXlsFps, vbMaximizedFocus
End Sub

Function StopXlsPsCxt$()
BfrClr
BfrV "PowerShell -Command ""try{Stop-Process -Id{try{(Get-Process -Name Excel).Id}finally{}.invoke()}"
BfrV "Pause..."
StopXlsPsCxt = BfrLines
End Function

Function StopXlsFps$()
StopXlsFps = TmpHom & "StopXls.Ps1"
End Function
