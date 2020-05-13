Attribute VB_Name = "MxIdeCacMth"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CNs$ = "Mth.Cac"
Const CMod$ = CLib & "MxIdeMthCac."
Public Const CacMdFF$ = MdFF

Function CacMthDrsP() As Drs
CacMthDrsP = CacMthDrszP(CPj)
End Function

Function CacMthDrszP(P As VBProject) As Drs
EnsCacMthDrs P
Dim O As Drs: Dim C As VBComponent: For Each C In P.VBComponents
    O = AddDrs(O, CacMthDrs(C.CodeModule))
Next
CacMthDrszP = O
End Function

Function CacMthDrsM() As Drs
Dim M As CodeModule: Set M = CMd
EnsCacMth M
CacMthDrsM = CacMthDrs(M)
End Function

Function CacMthDrs(M As CodeModule) As Drs
CacMthDrs = DrszFt(CacMthDrsFtzM(M))
End Function

Sub EnsCacMthDrs(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    EnsCacMth C.CodeModule
Next
End Sub

Private Sub IsCacMthOut__Tst()
Debug.Print IsCacMthOut(CMd)
End Sub

Function IsCacMthOut(M As CodeModule) As Boolean
Dim A$: A = CacMthMD5(M)
If A = "" Then IsCacMthOut = True: Exit Function
IsCacMthOut = MD5zM(M) <> A
End Function

Function CacMthMD5$(M As CodeModule)
Dim F$: F = CacMthMD5FtzM(M): If NoFfn(F) Then Exit Function
CacMthMD5 = MD5(LineszFt(F))
End Function

Function CacMthMD5M$()
CacMthMD5M = CacMthMD5(CMd)
End Function

Function IsCacMthOutM() As Boolean
IsCacMthOutM = IsCacMthOut(CMd)
End Function


Function MD5zM$(M As CodeModule)
MD5zM = MD5(Srcl(M))
End Function

Function MD5M$()
MD5M = MD5(SrclM)
End Function
