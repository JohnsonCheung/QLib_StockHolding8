Attribute VB_Name = "MxVbDtaS12Ln"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbDtaS12Ln."

Function LyzS12(A As S12, Optional Sep$ = " ") As String()
LyzS12 = FmtSq(SqzDy(Av(Av(A.S1, A.S2))))
End Function

Function TmlzS12$(A As S12, Optional Sep$ = " ")
TmlzS12 = TmlzAp(A.S1, A.S2)
End Function

Function S12zTml(Tml$) As S12
S12zTml = S12zTRst(TRstzLn(Tml))
End Function

Function S12zLn(Ln, Optional Sep$ = " ") As S12
S12zLn = Brk1(Ln, Sep)
End Function

Function S12y(Ly$()) As S12()
Dim S: For Each S In Itr(Ly)
    PushS12 S12y, S12zLn(S)
Next
End Function

Function LineszS12y$(A() As S12)
Dim J&, O$(): For J = 0 To S12UB(A)
     PushIAy O, LyzS12(A(J))
Next
LineszS12y = Jn(O, Chr(&H14))
End Function

Function LyzS12y(A() As S12) As String()
Dim J&: For J = 0 To S12UB(A)
    PushIAy LyzS12y, LyzS12(A(J))
Next
End Function
