Attribute VB_Name = "MxIdeCacDbMthOp"
Option Explicit
Option Compare Text
Const CNs$ = "CacMth"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthCacOp."

Sub ClrCacMthP()
ClrCacMthzP CPj
End Sub

Sub ClrCacMthzP(P As VBProject)
DltPthIfEmp CacMthPthzP(P)
End Sub

Private Sub EnsCacMth__Tst()
EnsCacMth CMd
End Sub

Function EnsCacMthM() As Boolean
EnsCacMthM = EnsCacMth(CMd)
End Function

Function EnsCacMthP() As Boolean
EnsCacMthP = EnsCacMthzP(CPj)
End Function

Function EnsCacMth(M As CodeModule) As Boolean
If IsCacMthOut(M) Then
    WrtCacMthMd5 M
    WrtCacMthDrs M
    EnsCacMth = True
End If
End Function

Function EnsCacMthzP(P As VBProject) As Boolean
EnsCacMthPthzP P
Dim C As VBComponent: For Each C In P.VBComponents
    If EnsCacMth(C.CodeModule) Then EnsCacMthzP = True
Next
End Function

Sub WrtCacMthMD5M(): WrtCacMthMd5 CMd: End Sub
Sub WrtCacMthDrsM(): WrtCacMthDrs CMd: End Sub

Sub WrtCacMthMd5(M As CodeModule)
Dim F$: F = CacMthMD5FtzM(M)
WrtStr MD5zM(M), F, OvrWrt:=True
End Sub

Sub WrtCacMthDrs(M As CodeModule)
WrtDrs SrtDrs(MthDrszM(M), "Mthn Mdy"), CacMthDrsFtzM(M)
End Sub
