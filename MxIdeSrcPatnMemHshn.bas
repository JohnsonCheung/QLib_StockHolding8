Attribute VB_Name = "MxIdeSrcPatnMemHshn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcMemn."
Public Const HshnFF$ = "Memn Hshn"

Function HshnRx() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("#(\w[\w:-]*)#", IsGlobal:=True)
Set HshnRx = X
End Function
Function Hshn$(S): Hshn = MchszR(S, HshnRx): End Function ' Return fst :Hshn or :Blnk

Private Sub HshnyP__Tst()
BrwAy HshnyP
End Sub

Function HshnyP() As String()
HshnyP = HshnyzP(CPj)
End Function
Function HshnyzP(P As VBProject) As String()
HshnyzP = Hshny(SrclzP(P))
End Function

Function Hshny(S$) As String() ' ret :Hshny from @S
'Hshny:Cml :NoSpcStr #Hash-Name-Ay# A NoSpcStr quoted by #
Hshny = MchsyzR(S, HshnRx)
End Function

Function HasHshn(S) As Boolean
HasHshn = HshnRx.Test(S)
End Function

Function HshnDrsP() As Drs: HshnDrsP = HshnDrs(SrcP): End Function
Function HshnDrs(Ly$()) As Drs
Dim ODy():
Dim L: For Each L In Itr(Ly)
    PushSomSi ODy, W1HshnDr(L)
Next
HshnDrs = DrszFF(HshnFF, ODy)
End Function

Private Function W1HshnDr(Ln) As String() ' Either EmpAv or Sy::[Memn Hshn]
Dim H$: H = Hshn(Ln): If H = "" Then Exit Function
W1HshnDr = Sy(Memn(Ln), H)
End Function
