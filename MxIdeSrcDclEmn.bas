Attribute VB_Name = "MxIdeSrcDclEmn"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclEmn."
Private Enum AA123: AA: End Enum
Function EnmSrczN(Dcl$(), Enmn) As String()
EnmSrczN = AwBix(Dcl, XBix(Dcl, Enmn))
End Function

Function EnmBei(Dcl$(), Enmn) As Bei
EnmBei = XBei(Dcl, XBix(Dcl, Enmn))
End Function
'-------------------------------
Private Sub EnmSrclAy__Tst()
BrwLinesy EnmSrclAy(DclP)
End Sub

Function EnmSrclAy(Dcl$()) As String()
Dim B() As Bei: B = W2Beiy(Dcl)
Dim J%: For J = 0 To BeiUB(B)
    PushI EnmSrclAy, JnCrLf(AwBei(Dcl, B(J)))
Next
End Function

Private Function W2Beiy(Dcl$()) As Bei()
Dim B: For Each B In Itr(W2Bixy(Dcl))
    PushBei W2Beiy, XBei(Dcl, B)
Next
End Function

Private Function W2Bixy(Dcl$()) As Integer()
Dim L, O%: For Each L In Itr(Dcl)
    If IsEnmln(L) Then PushI W2Bixy, O
    O = O + 1
Next
End Function
'-------------------------------------------------

Function EnmnyP() As String()
EnmnyP = EnmnyzP(CPj)
End Function

Function EnmnyM() As String()
EnmnyM = Enmny(DclM)
End Function

Function EnmnyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy EnmnyzP, XEnmnyzM(C.CodeModule)
Next
End Function
'------------------------------
Function Enmny(Dcl$()) As String()
Dim L: For Each L In Itr(Dcl)
   PushNB Enmny, XEnmn(L)
Next
End Function
'------------------------------
Function NEnm%(Src$())
Dim L, O%
For Each L In Itr(Src)
   If IsEnmln(L) Then O = O + 1
Next
NEnm = O
End Function
Function EnmSrcM(Enmn) As String()
EnmSrcM = EnmSrczM(CMd, Enmn)
End Function
Function EnmSrczM(M As CodeModule, Enmn) As String()
EnmSrczM = EnmSrc(Dcl(M), Enmn)
End Function

Function EnmSrc(Dcl$(), Enmn) As String()
EnmSrc = AwBei(Dcl, EnmBei(Dcl, Enmn))
End Function

Function NEnmzM%(M As CodeModule)
NEnmzM = NEnm(Dcl(M))
End Function

Function ShfTermEnm(OLin$) As Boolean
ShfTermEnm = ShfPfx(OLin, "Enum")
End Function

Sub BrwNDclLinP()
BrwDy NDclLinDy(CPj)
End Sub

Function IsEnmln(A) As Boolean
IsEnmln = HasPfx(RmvMdy(A), "Enum ")
End Function

'---===============================================
Private Function XBei(Dcl$(), Bix) As Bei
XBei = Bei(Bix, W1Eix(Dcl, Bix))
End Function

Private Function W1Eix%(Dcl$(), Bix)
W1Eix = VbItmEix(Dcl, Bix, "Enum")
End Function

Private Function XBix%(Dcl$(), Enmn)
Dim O%, L: For Each L In Itr(Dcl)
    If XEnmn(L) = Enmn Then
        XBix = O
        Exit Function
    End If
    O = O + 1
Next
XBix = -1
End Function

Private Function XEnmnyzM(M As CodeModule) As String()
XEnmnyzM = Enmny(Dcl(M))
End Function

Private Function XEnmn$(Ln)
Dim L$: L = RmvMdy(Ln)
If ShfPfx(L, "Enum ") Then XEnmn = TakNm(L)
End Function
