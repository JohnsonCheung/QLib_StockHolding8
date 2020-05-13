Attribute VB_Name = "MxCliMHDDb"
Option Explicit
Option Compare Text
Const CNs$ = "Mhd.Db"
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxCliMHDDb."
Function DutyDDb() As Database: Set DutyDDb = Db(DutyDtaFb): End Function
Function DutyPDb() As Database: Set DutyPDb = Db(DutyFba): End Function
