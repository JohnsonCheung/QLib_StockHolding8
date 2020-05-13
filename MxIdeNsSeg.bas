Attribute VB_Name = "MxIdeNsSeg"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeNsSeg."

Private Sub NsSegAyP__Tst()
Vc NsSegAyP
End Sub

Function NsSegAyP() As String()
NsSegAyP = NsSegAyzP(CPj)
End Function

Function NsSegAyzP(P As VBProject) As String()
Dim N: For Each N In NsAyzP(P)
    PushNDupAy NsSegAyzP, SplitDot(N)
Next
End Function

Function MdNyWhNsSegssP(NsSegss$) As String()
MdNyWhNsSegssP = MdNyWhNsSegsszP(CPj, NsSegss)
End Function

Function MdNyWhNsSegsszP(P As VBProject, NsSegss$) As String()
Dim NsGp(), Mdny$()
    Dim C As VBComponent: For Each C In P.VBComponents
        Dim N$: N = CNsv(Dcl(C.CodeModule))
        If N <> "" Then
            PushI Mdny, C.Name
            PushI NsGp, SplitDot(N)
        End If
    Next
Dim Ixy&(): Ixy = IxyzTagGp(NsGp, SyzSS(NsSegss))
MdNyWhNsSegsszP = AwIxy(Mdny, Ixy)
End Function

Function IxyzTagGp(TagGp(), WhTagAy$()) As Long()
Dim TagAy$()
Dim J&: For J = 0 To UB(TagGp)
    TagAy = TagGp(J)
    If IsSuperAy(TagAy, WhTagAy) Then PushI IxyzTagGp, J
Next
End Function
