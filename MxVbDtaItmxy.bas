Attribute VB_Name = "MxVbDtaItmxy"
Option Explicit
Option Compare Text
Type Itmxy: Itm As String: Ixy() As Long: End Type 'Deriving(Ctor Ay)

Function DisIxyzItmxyAy(I() As Itmxy) As Long()
Dim J%: For J = 0 To ItmxyUB(I)
    PushNDupAy DisIxyzItmxyAy, I(J).Ixy
Next
End Function
Function Itmxy(Itm, Ixy&()) As Itmxy
With Itmxy
    .Itm = Itm
    .Ixy = Ixy
End With
End Function
Function AddItmxy(A As Itmxy, B As Itmxy) As Itmxy(): PushItmxy AddItmxy, A: PushItmxy AddItmxy, B: End Function
Sub PushItmxyAy(O() As Itmxy, A() As Itmxy): Dim J&: For J = 0 To ItmxyUB(A): PushItmxy O, A(J): Next: End Sub
Sub PushItmxy(O() As Itmxy, M As Itmxy): Dim N&: N = ItmxySi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function ItmxySi&(A() As Itmxy): On Error Resume Next: ItmxySi = UBound(A) + 1: End Function
Function ItmxyUB&(A() As Itmxy): ItmxyUB = ItmxySi(A) - 1: End Function
