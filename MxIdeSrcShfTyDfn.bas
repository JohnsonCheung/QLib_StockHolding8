Attribute VB_Name = "MxIdeSrcShfTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcShfTyDfn."

Function ShfTyDfnNm$(OLin$)
Dim A$: A = T1(OLin)
If IsTyDfnNm(A) Then
    ShfTyDfnNm = A
    OLin = RmvT1(OLin)
End If
End Function

Function ShfColonTy$(OLin$)
':ColonTy: :Str ! #Colon-Type# it is a Term with fst chr is : and rest is [DfnTyNm]
Dim A$: A = T1(OLin)
If IsDfnTy(A) Then
    ShfColonTy = A
    OLin = RmvT1(OLin)
End If
End Function

Function ShfMemNm$(OLin$)
Dim A$: A = T1(OLin)
If IsMemNm(A) Then
    ShfMemNm = A
    OLin = RmvT1(OLin)
End If
End Function
