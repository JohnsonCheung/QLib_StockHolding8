Attribute VB_Name = "MxDtaEr"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaEr."

Function EoDupCol(D As Drs, C$) As String()
Dim B As Drs: B = DwDup(D, C)
Dim Msg$: Msg = "Dup [" & C & "]"
EoDupCol = EoMsgDrs(Msg, B)
End Function
