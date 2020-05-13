Attribute VB_Name = "MxXlsAddIn"
Option Compare Text
Option Explicit
Const CNs$ = "Xls.AddIn"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsAddIn."

Function AddinDrs(A As Excel.Application) As Drs
AddinDrs = DrszItp(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function

Sub DmpAddinX()
DmpAddinzX Xls
End Sub

Sub DmpCAddin()
DmpAddinzX Xls
End Sub
Sub DmpAddinzX(X As Excel.Application)
DmpDrsR AddinDrs(X)
End Sub

Function AddinWs(X As Excel.Application) As Worksheet
Set AddinWs = VisWs(WszDrs(AddinDrs(X)))
End Function

Function Addin(X As Excel.Application, FxaNm) As Excel.Addin
Dim I As Excel.Addin
For Each I In X.AddIns
    If I.Name = FxaNm & ".xlam" Then Set Addin = I
Next
End Function
