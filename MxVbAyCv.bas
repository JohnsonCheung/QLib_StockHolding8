Attribute VB_Name = "MxVbAyCv"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Map"
Const CMod$ = CLib & "MxVbAyCv."

Function CvSy(A) As String(): CvSy = A: End Function
Function CvObj(A) As Object: Set CvObj = A: End Function
Function CvBytAy(A) As Byte(): CvBytAy = A: End Function
Function CvIntAy(A) As Integer(): CvIntAy = A: End Function
Function CvLngAy(A) As Long(): CvLngAy = A: End Function

Function CvAv(V) As Variant() 'Ret Av if @V is Av or Empty, else thw error
Dim T As VbVarType: T = VarType(V)
Select Case True
Case T = vbArray + vbVariant Or T = vbEmpty
    If Si(V) = 0 Then Exit Function
    CvAv = V
    Exit Function
End Select
Thw CSub, "Givan V must be vbArray+vbVariant or vbEmpty", "TypeName(V)", TypeName(V)
End Function

Function SyzV(Str_or_Sy_or_Ay_or_EmpMis_or_Oth) As String()
Dim A: A = Str_or_Sy_or_Ay_or_EmpMis_or_Oth
Select Case True
Case IsStr(A): PushI SyzV, A
Case IsSy(A): SyzV = A
Case IsArray(A): SyzV = SyzAy(A)
Case IsEmpty(A) Or IsMissing(A)
Case Else: SyzV = Sy(A)
End Select
End Function
Function CvStr$(V)
Select Case True
Case IsNull(V):
Case IsArray(V): CvStr = "*" & TypeName(V) & "[" & UB(V) & "]"
Case Else: CvStr = V
End Select
End Function
