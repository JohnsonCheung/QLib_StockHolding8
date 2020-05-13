Attribute VB_Name = "MxIdeDef"
Option Explicit
Option Compare Text
Private Sub DefnRx__Tst()
BrwAy MchsyzR(SrclP, DefnRx)
End Sub

Function DefnRx() As RegExp
'Def:Cml #Defintion# Definition-Here
'Defn:Cml :Nm #Definition-Name#
Static X As RegExp: If IsNothing(X) Then Set X = Rx("'([A-Za-z][\w-]+)\:\:", IsGlobal:=True)
Set DefnRx = X
End Function
