Attribute VB_Name = "MxXlsWbOp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWbOp."

Sub SavWbQuit(Wb As Workbook)
Dim X As Excel.Application
Set X = Wb.Application
Wb.Close True
X.Quit
End Sub

Sub VArrangeWb(X As Excel.Application)
Dim W As Excel.Window: For Each W In X.Windows
    W.Activate
    W.WindowState = xlNormal
    W.Visible = True
Next
X.Windows.Arrange xlArrangeStyleVertical
End Sub
