Attribute VB_Name = "MxVbDtaTag"
Option Explicit
Option Compare Text
Type Tagu: Key As String: Tag() As String: End Type ' Deriving(Ay Ctor)
Function KyzTagss(T() As Tagu, Tagss$) As String(): KyzTagss = KyzTagy(T, SyzSS(Tagss)): End Function
Function KyzTagy(T() As Tagu, Tagy$()) As String()
Dim J%: For J = 0 To TaguUB(T)
    If HasIntersect(T(J).Tag, Tagy) Then PushI KyzTagy, T(J).Key
Next
End Function
Function TaguUB&(A() As Tagu)
End Function
