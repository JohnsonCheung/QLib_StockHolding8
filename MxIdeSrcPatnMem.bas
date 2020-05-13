Attribute VB_Name = "MxIdeSrcPatnMem"
Option Explicit
Option Compare Text
Enum eHshOpt: eExlHsh: eInlHsh: End Enum
'**Memn
Private Sub Memny__Tst(): BrwAy Memny(SrclP): End Sub
Function HasMemn(S) As Boolean: HasMemn = HasRx(S, MemnRx): End Function
Function Memn$(S): Memn = MchszR(S, MemnRx): End Function ' #memonic-name# a name between 2 hashChr
Function MemnRx() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("#([A-Za-z][\w\.-]*)#", IsGlobal:=True)
Set MemnRx = X
End Function
Function Memny(S$, Optional H As eHshOpt) As String()
If H = eExlHsh Then
    Memny = SMchsyzR(S, MemnRx)
Else
    Memny = MchsyzR(S, MemnRx)
End If
End Function
