Attribute VB_Name = "MxIdeMthnMdn"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeMthnMdn."

Function MdnyzPubMth(P As VBProject, PubMthn) As String()
Dim I, M As CodeModule: For Each I In ModItrzP(P)
    Set M = I
    If HasPubMth(Src(M), PubMthn) Then PushI MdnyzPubMth, Mdn(M)
Next
End Function

Private Sub MdnyzPubMth__Tst()
Dim P As VBProject, PubMthn
GoSub Z
Exit Sub
Z:
    D MdnyzPubMth(CPj, "AA")
    Stop
    Return
End Sub

Function MdnyzPubMthP(PubMthn) As String()
MdnyzPubMthP = MdnyzPubMth(CPj, PubMthn)
End Function
