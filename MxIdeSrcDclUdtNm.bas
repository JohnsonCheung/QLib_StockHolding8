Attribute VB_Name = "MxIdeSrcDclUdtNm"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclUdtNm."

Function Udtn$(Ln)
Dim L$: L = RmvMdy(Ln)
If Not ShfTermX(L, "Type") Then Exit Function
Udtn = TakNm(L)
End Function

Function UdtnzL$(Ln)
UdtnzL = Udtn(Ln)
End Function

Function UdtMzN(Udtn$) As Udt: UdtMzN = UdtzN(DclM, Udtn): End Function ' #Fst-Udt-In-CMd#
Function UdtPzN(Udtn$) As Udt '#Fst-Udt-In-CPj#
Dim C As VBComponent: For Each C In CPj.VBComponents
    UdtPzN = UdtzN(Dcl(C.CodeModule), Udtn)
    If UdtPzN.Udtn <> "" Then Exit Function
Next
End Function

'**Udtn
Private Sub Udtn__Tst()
Debug.Assert Udtn("Type Udt") = "Udt"
Debug.Assert Udtn("Private Type Udt") = "Udt"
End Sub

Function UdtnyM() As String()
UdtnyM = UdtnyzM(CMd)
End Function

Function UdtnyzM(M As CodeModule) As String()
UdtnyzM = Udtny(Dcl(M))
End Function

Function UdtnyP() As String()
UdtnyP = UdtnyzP(CPj)
End Function

Function UdtnyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy UdtnyzP, UdtnyzM(C.CodeModule)
Next
End Function

Function Udtny(Dcl$()) As String()
Dim L: For Each L In Itr(Dcl)
    PushNB Udtny, Udtn(L)
Next
End Function

Function HasUdtn(Dcl$(), Udtn$) As Boolean
Dim L: For Each L In Itr(Dcl)
    If UdtnzL(L) = Udtn Then HasUdtn = True: Exit Function
Next
End Function
