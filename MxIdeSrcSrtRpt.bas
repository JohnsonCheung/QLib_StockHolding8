Attribute VB_Name = "MxIdeSrcSrtRpt"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcSrtRpt."
Function SrtRpt(Src$()) As String()
Dim X As Dictionary
Dim Y As Dictionary
Set X = MthlDic(Src)
Set Y = SrtDic(X)
SrtRpt = FmtCprDic(X, Y, "BefSrt", "AftSrt")
End Function

Private Sub SrtRpt__Tst()
Brw SrtRptzM(CMd)
End Sub

Property Get SrtRptM() As String()
SrtRptM = SrtRptzM(CMd)
End Property

Function SrtSrc(Src$()) As String()
Dim A$(): A = SVy(SrtDic(MthlDic(Src)))
SrtSrc = SplitCrLf(Jn(A, vb2CrLf))
End Function

Function SrtRptzP(P As VBProject) As String()
Dim O$(), C As VBComponent
For Each C In P.VBComponents
    PushIAy O, SrtRptzM(C.CodeModule)
Next
SrtRptzP = O
End Function

Function SrtRptDiczP(P As VBProject) As Dictionary
Dim C As VBComponent, O As New Dictionary, Md As CodeModule
    For Each C In P.VBComponents
        Set Md = C.CodeModule
        O.Add Mdn(Md), SrtRptzM(Md)
    Next
Set SrtRptDiczP = O
End Function

Function SrtRptzM(M As CodeModule) As String()
SrtRptzM = SrtRpt(Src(M))
End Function
