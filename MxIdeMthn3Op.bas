Attribute VB_Name = "MxIdeMthn3Op"
Option Explicit
Option Compare Text
'** Where

Function Mthn3yWhInlPrv(N() As Mthn3, InlPrv As Boolean) As Mthn3()
If InlPrv Then Mthn3yWhInlPrv = N: Exit Function
Dim J&: For J = 0 To Mthn3UB(N)
    If N(J).ShtMdy <> "Prv" Then PushMthn3 Mthn3yWhInlPrv, N(J)
Next
End Function
Function Mthnyz3(N() As Mthn3) As String()
Dim J&: For J = 0 To Mthn3UB(N)
    PushI Mthnyz3, N(J).Nm
Next
End Function

