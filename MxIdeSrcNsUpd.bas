Attribute VB_Name = "MxIdeSrcNsUpd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcNsUpd."
Sub UpdNsP()
UpdNszP CPj
End Sub

Private Sub UpdNszP(P As VBProject)
Dim L: For Each L In LyzFt(MdNsFt(P))
    Dim A As S12: A = Brk1(L, " ")
    If HasMd(P, A.S1) Then
        EnsCNs Md(A.S1), A.S2
    End If
Next
End Sub

Sub EdtNsP()
EdtNs CPj
End Sub

Sub EdtNs(P As VBProject)
Dim F$: F = MdNsFt(P)
WrtAy AliLyz1T(MdNsLy(P)), F, OvrWrt:=True
VcFt F
End Sub

Function MdNsLyP() As String()
MdNsLyP = MdNsLy(CPj)
End Function

Function MdNsLy(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI MdNsLy, C.Name & " " & CNsv(Dcl(C.CodeModule))
Next
End Function
