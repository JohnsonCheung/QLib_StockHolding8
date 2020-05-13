Attribute VB_Name = "MxIdeMdn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdn."

Function MdGnRx() As RegExp
Static X As RegExp
'If IsNothing(X) Then Set X = Rx("Mx[A-Z][a-z]+[A-Z][a-z0-9]+")
If IsNothing(X) Then Set X = Rx("^Mx[A-Z][a-z]+")
Set MdGnRx = X
End Function

Function MdGnyP() As String()
MdGnyP = MdGnyzP(CPj)
End Function

Function MdGnyzP(P As VBProject) As String()
MdGnyzP = QSrt(MdGny(Mdny(P)))
End Function

Function MdGny(Mdny$()) As String()
MdGny = AwRx(Mdny, MdGnRx)
End Function

Function MdGn$(Mdn)
MdGn = Mchs(Mdn, MdGnRx)
End Function

Function MdNyP() As String()
MdNyP = Mdny(CPj)
End Function

Function MdnyBySubStr(MdnSubStr$, Optional C As eCas) As String()
MdnyBySubStr = AwSubStr(Mdny(CPj), MdnSubStr, C)
End Function

Function MdnyzPfx(MdnPfx) As String()
MdnyzPfx = AwPfx(Mdny(CPj), MdnPfx)
End Function

Function Mdny(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMd(C) Then
        PushI Mdny, C.Name
    End If
Next
End Function

Function MdNyzV(A As Vbe) As String()
Dim P As VBProject: For Each P In A.VBProjects
    PushIAy MdNyzV, Mdny(P)
Next
End Function

Function ModNyP() As String()
ModNyP = ModNy(CPj)
End Function

Function CClsNy() As String()
CClsNy = ClsNy(CPj)
End Function

Function ModNy(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMod(C) Then PushI ModNy, C.Name
Next
End Function

Private Sub ClsNy__Tst()
DmpAy ClsNy(CPj)
End Sub

Private Sub MdAy__Tst()
Dim O() As CodeModule
O = MdAyzP(CPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print Mdn(Md)
Next
End Sub

Private Sub MdzPjny__Tst()
'DmpAy MdzPjny(CPj)
End Sub
