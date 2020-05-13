Attribute VB_Name = "MxXlsTy"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxXlsTy."
Enum EmXlsTy
    EiNum ' N
    EiTxt ' T
    EiTorN 'TorN
    EiDte 'Dte
    EiBool 'B
End Enum
Function XlsTyAyzCsv(XlsTyCsv$) As EmXlsTy()
Dim Ay$(): Ay = Split(XlsTyCsv, ",")
Dim J%: For J = 0 To UBound(Ay)
    Ay(J) = Trim(Ay(J))
    Select Case Ay(J)
    Case "T": PushI XlsTyAyzCsv, EiTxt
    Case "N": PushI XlsTyAyzCsv, EiNum
    Case "TorN": PushI XlsTyAyzCsv, EiTorN
    Case "Dte": PushI XlsTyAyzCsv, EiDte
    Case "B": PushI XlsTyAyzCsv, EiBool
    Case Else: Thw "XlsTyAyzCsv", "XlsTyStr should be T or N or TorN or Dte or B, but now[" & Ay(J) & "]", "XlsTyCsv", XlsTyCsv
    End Select
Next
End Function

Function IsEqXlsTy(XlsTy As EmXlsTy, AdoTy As ADODB.DataTypeEnum) As Boolean
Select Case True
Case XlsTy = EiBool: IsEqXlsTy = AdoTy = adBoolean
Case XlsTy = EiDte:  IsEqXlsTy = AdoTy = adDate
Case XlsTy = EiNum:  IsEqXlsTy = AdoTy = adDouble
Case XlsTy = EiTxt:  IsEqXlsTy = AdoTy = adVarWChar
Case XlsTy = EiTorN: IsEqXlsTy = (AdoTy = adVarWChar) Or (AdoTy = adDouble)
End Select
End Function

Function ShtXlsTy$(A As EmXlsTy)
Dim O$
Select Case True
Case A = EiNum: O = "Num"
Case A = EiTxt: O = "Txt"
Case A = EiTorN: O = "Txt or Num"
Case A = EiBool: O = "Bool"
Case A = EiDte: O = "Dte"
End Select
ShtXlsTy = O
End Function
