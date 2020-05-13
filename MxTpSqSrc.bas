Attribute VB_Name = "MxTpSqSrc"
Option Explicit
Option Compare Text
Type Stmtll: A As String: End Type
Type Swl: Swn As String: Op As String: Termy() As String: End Type ' Deriving(Ctor Ay)
Type SqSrc
    Sw() As Swl
    Pm() As String
    Stmtlly() As Stmtll
End Type
Function SqSrczT(SqTp$()) As SqSrc
'With SqTpSrczT
'    .Oth = W1Oth
'    .Pm = W1Oth
'    .Rmk = Sy()
'    .Sq = W1Oth
'    .Sw = W1Oth
'End With
End Function

Private Function W1Oth() As LLn()

End Function

Function Swl(Swn, Op, Termy$()) As Swl
With Swl
    .Swn = Swn
    .Op = Op
    .Termy = Termy
End With
End Function
Function AddSwl(A As Swl, B As Swl) As Swl(): PushSwl AddSwl, A: PushSwl AddSwl, B: End Function
Sub PushSwlAy(O() As Swl, A() As Swl): Dim J&: For J = 0 To SwlUB(A): PushSwl O, A(J): Next: End Sub
Sub PushSwl(O() As Swl, M As Swl): Dim N&: N = SwlSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SwlSi&(A() As Swl): On Error Resume Next: SwlSi = UBound(A) + 1: End Function
Function SwlUB&(A() As Swl): SwlUB = SwlSi(A) - 1: End Function
