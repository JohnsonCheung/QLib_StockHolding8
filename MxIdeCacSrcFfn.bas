Attribute VB_Name = "MxIdeCacSrcFfn"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcFfnCac."

Function IsSrcFfnCac(C As VBComponent) As Boolean
IsSrcFfnCac = IsEqFfn(SrcFfn(C), Src(C.CodeModule), eCprTimSi)
End Function
