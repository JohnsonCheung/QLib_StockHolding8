Attribute VB_Name = "MxIdeSrcPatnDefn"
Option Explicit
Option Compare Text
Enum eDefrOpt: eNoColonQ: eWiColonQ: End Enum
#If Doc Then
'Defr:Cml #Definition-Reference# :Patn
#End If

'**Def
Private Sub Defty__Tst(): VcAy Defty(SrclP, eWiColonQ): End Sub ':Johnson
Function HasDefr(S) As Boolean: HasDefr = HasRx(S, DefrRx): End Function
Function Defn$(S): Defn = MchszR(S, DefrRx): End Function ' #Def-name# a name between 2 hashChr
Function DefrRx() As RegExp '#Defintion-Reference# /:xxx / or /:xxx$/
Static X As RegExp: If IsNothing(X) Then Set X = Rx(":([A-Za-z][\w\.-]*$)", IsGlobal:=True, MultiLine:=True)
'Static X As RegExp: If IsNothing(X) Then Set X = Rx(":([A-Za-z][\w\.-]*) |:([A-Za-z][\w\.-]*)$", IsGlobal:=True)
Set DefrRx = X
End Function
Function Defty(S$, Optional H As eDefrOpt) As String()
If H = eExlHsh Then
    Defty = SMchsyzR(S, DefrRx)
Else
    Defty = MchsyzR(S, DefrRx)
End If
End Function

Function Cmlln$(S) ' #Camel-definition-name# fmt:: 'xxx:

End Function
