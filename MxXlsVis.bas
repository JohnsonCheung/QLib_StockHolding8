Attribute VB_Name = "MxXlsVis"
Option Explicit
Option Compare Text
Const CNs$ = "Xls.Op"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsVis."
Private A() As Boolean
Sub PushXlsVis()
PushI A, Xls.Visible
End Sub
Sub PopXlsVis()
Xls.Visible = PopI(A)
End Sub
Sub PushXlsVisHid()
PushXlsVis
Xls.Visible = False
End Sub

Function VisWb(A As Workbook) As Workbook
VisXls A.Application
Set VisWb = A
End Function

Function VisCXls() As Excel.Application
Set VisCXls = VisXls(Xls)
End Function

Function VisXls(A As Excel.Application) As Excel.Application
If Not A.Visible Then A.Visible = True
Set VisXls = A
End Function

Function VisLo(A As ListObject) As ListObject
VisXls A.Application
Set VisLo = A
End Function

Function VisWs(S As Worksheet) As Worksheet
VisXls S.Application
Set VisWs = S
End Function

Function VisRg(R As Range) As Range
VisXls R.Application
Set VisRg = R
End Function
