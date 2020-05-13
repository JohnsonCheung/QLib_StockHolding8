Attribute VB_Name = "MxIdeMthnRel"
Option Explicit
Option Compare Text
Const CNs$ = "Mthn"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthnRel."

Function PubMthnRelMdNyP() As Dictionary
Set PubMthnRelMdNyP = PubMthnRelMdnzP(CPj)
End Function

Function MthnRelCmlV() As Dictionary
Set MthnRelCmlV = MthnRelCmtzV(CVbe)
End Function

Function MthnRelCmtzV(A As Vbe) As Dictionary
Dim O As New Dictionary, I
For Each I In MthnyzV(A)
    PushRelLin O, Cmlss(I)
Next
Set MthnRelCmtzV = O
End Function

Private Sub PubMthnRelMdnzP__Tst()
BrwRel PubMthnRelMdnzP(CPj)
End Sub

Function PubMthnRelMdnzP(P As VBProject) As Dictionary
Dim O As New Dictionary, Mthn, Mdn$
Dim C As VBComponent: For Each C In P.VBComponents
    Mdn = C.Name
    For Each Mthn In Itr(PubMthnyzS(Src(C.CodeModule)))
        PushParChd O, Mthn, Mdn
    Next
Next
Set PubMthnRelMdnzP = O
End Function

Function MthnRelMdnzP(P As VBProject) As Dictionary
Dim O As New Dictionary, Mthn, Mdn$
Dim C As VBComponent: For Each C In P.VBComponents
    Mdn = C.Name
    For Each Mthn In Itr(Mthny(Src(C.CodeModule)))
        PushParChd O, Mthn, Mdn
    Next
Next
Set MthnRelMdnzP = O
End Function

Function MthnRelMdnP() As Dictionary
Set MthnRelMdnP = MthnRelMdnzP(CPj)
End Function
