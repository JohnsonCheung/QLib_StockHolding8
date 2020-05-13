Attribute VB_Name = "JMxApp"
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "JMxApp."
#If False Then
Option Compare Text
Option Explicit
Option Base 0

Sub CpyIFxToIW(IFxPmn$)
CpyFfnIfDif Pmv(IFxPmn), WIFx(IFxPmn)
End Sub

Function WIFx$(IFxPmn$)
WIFx = Fn(Pmv(IFxPmn))
End Function

#End If
