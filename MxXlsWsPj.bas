Attribute VB_Name = "MxXlsWsPj"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxXlsWsPj."

Function MdzWs(S As Worksheet) As CodeModule
Set MdzWs = CmpzWs(S).CodeModule
End Function

Function CmpzWs(S As Worksheet) As VBComponent
Set CmpzWs = FstObjByNm(PjzWs(S).VBComponents, S.CodeName)
End Function

Function PjzWb(B As Workbook) As VBProject
Set PjzWb = B.VBProject
End Function

Function PjzWs(S As Worksheet) As VBProject
Set PjzWs = WbzWs(S).VBProject
End Function

Function PjzRg(A As Range) As VBProject
Set PjzRg = WbzRg(A).VBProject
End Function
