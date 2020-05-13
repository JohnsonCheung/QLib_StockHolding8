Attribute VB_Name = "MxIdeSrcGenDeriUdtOp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcGenDeriUdtOp."
Private Sub DeriUdt__Tst(): DeriUdtzM CMd: End Sub
Private Sub W1HasDeriUdt__Tst()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If W1HasDeriUdt(C.CodeModule) Then Debug.Print C.Name
Next
End Sub

Private Sub AA_DeriUdt(): End Sub
Sub DeriUdtM(): DeriUdtzM CMd: End Sub
Sub DeriUdtzMdn(Mdn): DeriUdtzM Md(Mdn): End Sub
Sub DeriUdtzM(M As CodeModule)
ChkPjSav PjzM(M)
Dim U() As Udt: U = UdtyzM(M)
Dim J%: For J = 0 To UdtUB(U)
    VVDeriUdt M, U(J)
Next
End Sub
Sub DeriUdtzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    If W1HasDeriUdt(C.CodeModule) Then
        DeriUdtzM C.CodeModule
    End If
Next
End Sub
Private Function W1HasDeriUdt(M As CodeModule) As Boolean
W1HasDeriUdt = W1Rx.Test(Dcll(M))
End Function
Private Function W1Rx() As RegExp
Static R As RegExp: If IsNothing(R) Then Set R = Rx("Deriving\(.+\)", IsGlobal:=True)
Set W1Rx = R
End Function

Private Sub VVDeriUdt(M As CodeModule, U As Udt) ' Derive user defined type of given module by Udt
Dim C As UdtDvc: C = UdtDvc(U)
Dim N$: N = U.Udtn
EnsUdt M, C.OptUdtl, N
EnsMth M, C.Mthl, C.Mthny
End Sub

