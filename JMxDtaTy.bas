Attribute VB_Name = "JMxDtaTy"
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "JMxDtaTy."
#If False Then
Option Explicit
Function DaoTyAyzTyLetr(TyLetrCsv$) As DAO.DataTypeEnum()
Dim Ay$(): Ay = Split(TyLetrCsv, ",")
Dim J%: For J = 0 To UBound(Ay)
    Ay(J) = Trim(Ay(J))
    Select Case Ay(J)
    Case "T": PushI DaoTyAyzTyLetr, dbText
    Case "N": PushI DaoTyAyzTyLetr, dbDouble
    Case Else: PmEr "DaoTyAyzTyLetr: TyLetr should be T or N. Now[" & Ay(J) & "]"
    End Select
Next
End Function

Function ShtDaoTy$(A As DAO.DataTypeEnum)
Dim O$
Select Case True
Case A = dbAttachment: O = "Att"
Case A = dbBigInt: O = "BigInt"
Case A = dbBinary: O = "Bin"
Case A = dbBoolean: O = "Bool"
Case A = dbByte: O = "Byt"
Case A = dbChar: O = "Chr"
Case A = dbCurrency: O = "Ccy"
Case A = dbDate: O = "Dte"
Case A = dbDecimal: O = "Dec"
Case A = dbDouble: O = "Dbl"
Case A = dbFloat: O = "Float"
Case A = dbInteger: O = "Int"
Case A = dbLong: O = "Lng"
Case A = dbMemo: O = "Mem"
Case A = dbSingle: O = "Sng"
Case A = dbText: O = "Txt"
Case A = dbTime: O = "Tim"
Case A = dbTimeStamp: O = "Stmp"
Case Else: ThwPgmEr "ShtDaoTy: Unsuppoert DaoTy[" & A & "]"
End Select
ShtDaoTy = O
End Function
#End If
