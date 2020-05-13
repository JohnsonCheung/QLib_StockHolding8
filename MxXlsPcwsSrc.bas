Attribute VB_Name = "MxXlsPcwsSrc"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsPcwsSrc."
Sub Worksheet_SelectionChange(ByVal Target As Range)
PutPcwsChd Target
End Sub
