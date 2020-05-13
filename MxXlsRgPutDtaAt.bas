Attribute VB_Name = "MxXlsRgPutDtaAt"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRgPutDtaAt."

Function PutDbtAt(Db As Database, T, At As Range) As Range
Set PutDbtAt = RgzSq(SqzT(Db, T), At)
End Function
