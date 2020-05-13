Attribute VB_Name = "MxIdeSrcQee"
Option Explicit
Option Compare Text
Private Sub Qeey__Tst(): BrwAy QeeyP: End Sub
Function QeeyP() As String(): QeeyP = Qeey(SrcP): End Function
Function Qeey(Src$()) As String() '#Quo-Eq-Eq-Ay# 1 SngQuo 2 Eq-Sign Array
Dim I: For Each I In Itr(Src)
    If IsQee(I) Then PushI Qeey, I
Next
End Function
Function IsQee(Ln) As Boolean: IsQee = HasPfx(Ln, "'=="): End Function ' #Is-Quo-Eq-Eq# is the @Ln has pfx '==

Private Sub Qssy__Tst(): BrwAy QssyP: End Sub
Function QssyP() As String(): QssyP = Qssy(SrcP): End Function
Function Qssy(Src$()) As String() '#Quo-Star-Star-Ay# 1 SngQuo 2 Star-Sign Array
Dim I: For Each I In Itr(Src)
    If IsQss(I) Then PushI Qssy, I
Next
End Function
Function IsQss(Ln) As Boolean: IsQss = HasPfx(Ln, "'**"): End Function ' #Is-Quo-Star-Star# is the @Ln has pfx '==

