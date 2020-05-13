Attribute VB_Name = "MxVbAySft"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbAySft."

Function ShfEle(OAy, Ele) As Boolean
Dim At&
Dim I: For Each I In Itr(OAy)
    If I = Ele Then
        ShfEle = True
        OAy = AeAt(OAy, At)
        Exit Function
    End If
    At = At + 1
Next
End Function
Function ShfIntBet%(OAy, A%, B%)
Dim At&
Dim I: For Each I In Itr(OAy)
    If IsNumeric(I) Then
        If IsBet(CInt(I), A, B) Then
            ShfIntBet = I
            
            OAy = AeAt(OAy, At)
            Exit Function
        End If
    End If
    At = At + 1
Next
End Function
