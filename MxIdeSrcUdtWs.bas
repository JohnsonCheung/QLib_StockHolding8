Attribute VB_Name = "MxIdeSrcUdtWs"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxIdeSrcUdtWs."

Function UdtWsP() As Worksheet
Set UdtWsP = UdtWszP(CPj)
End Function

Function UdtWszP(P As VBProject) As Worksheet
Set UdtWszP = FmtUdtWs(WszDrs(UdtDrszP(P)))
End Function

Function FmtUdtWs(UdtWs As Worksheet) As Worksheet
Set FmtUdtWs = UdtWs
End Function

Function UdtDrszP(P As VBProject) As Drs

End Function
