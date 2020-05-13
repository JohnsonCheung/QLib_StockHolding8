Attribute VB_Name = "MxVbDtaTRst"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "MxVbDtaTRst."
Type TRst
    T As String
    Rst As String
End Type
Sub PushTRst(O() As TRst, M As TRst)
Dim N&: N = TRstSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function TRstWhEmpRst(A() As TRst) As TRst()
Dim J&: For J = 0 To TRstUB(A)
    If A(J).Rst <> "" Then
        PushTRst TRstWhEmpRst, A(J)
    End If
Next
End Function
