Attribute VB_Name = "MxXlsMaxi"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsMaxi."

Sub MaxvXls(X As Excel.Application)
MaxvXls X
VisXls X
End Sub

Function MaxiXls(X As Excel.Application) As Excel.Application
X.WindowState = xlMinimum
X.Visible = True
Set MaxiXls = X
End Function
Function MiniXls(X As Excel.Application) As Excel.Application
X.WindowState = xlMinimized
X.Visible = True
Set MiniXls = X
End Function

Sub MinvWb(Wb As Workbook)
MinvXls Wb.Application
End Sub

Sub MinvRg(R As Range)
MinvXls R.Application
End Sub

Sub MiniWb(Wb As Workbook)
MiniXls Wb.Application
End Sub

Function MinvXls(X As Excel.Application) As Excel.Application
X.WindowState = xlNormal
X.Visible = True
X.Left = 0
X.Top = 0
X.Width = 1
X.Height = 1
Set MinvXls = X
End Function

#If False Then
Sub MaxiXls(X As Excel.Application): X.Visible = True: X.WindowState = xlMaximized: End Sub
Sub MaxiWs(Ws As Worksheet):         MaxiXls Ws.Application: End Sub

Sub MinvXls(X As Excel.Application)
X.Visible = True
X.WindowState = xlNormal
X.Left = 0
X.Top = 0
X.Width = 1
X.Height = 1
End Sub
Sub MinvWb(Wb As Workbook):          MinvXls Wb.Application: End Sub
Sub MinvRg(Rg As Range):             MinvXls Rg.Application: End Sub
Sub MiniXls(X As Excel.Application): X.Visible = True: X.WindowState = xlMinimized: End Sub
Sub MiniWb(Wb As Workbook): MiniXls Wb.Application: End Sub

#End If
