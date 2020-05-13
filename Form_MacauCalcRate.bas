VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MacauCalcRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_MacauCalcRate."
Private Sub Cmd_MacauOverrideRate_Click()
DoCmd.OpenTable "MacauOverrideRate"
End Sub
