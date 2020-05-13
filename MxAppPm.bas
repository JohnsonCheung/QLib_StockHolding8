Attribute VB_Name = "MxAppPm"
Option Compare Text
Option Explicit
Const CNs$ = "App.Pm"
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxAppPm."
Sub OpnIPth(): BrwPth AppIPth: End Sub
Sub OpnOPth(): BrwPth AppOPth: End Sub
Sub OpnTpPth(): BrwPth TpPthP: End Sub

Function AppIPth$():  AppIPth = CDbPth & "SAPDownloadExcel\": End Function
Function AppIWPth$(): AppIWPth = CDbPth & "SAPDownloadExcel\Wrk\": End Function
Function AppOPth$():  AppOPth = CDbPth & "Output\": End Function
Function TpPthP$()
TpPthP = TpPthzP(CPj)
End Function
Function TpPthzP$(P As VBProject)
TpPthzP = EnsPth(AssPthzP(P) & "Templates\")
End Function
