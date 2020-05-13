Attribute VB_Name = "MxCliMHD"
Option Compare Text
Option Explicit
Const CLib$ = "QClient."
Const CNs$ = "Client.Mhd"
Const CMod$ = CLib & "MxCliMHD."
Function DutyDtaDb() As Database: Set DutyDtaDb = Db(DutyDtaFb): End Function
Function StkHld8PgmDb() As Database: Set StkHld8PgmDb = Db(StkHld8Fba): End Function
Function StkHld8Db() As Database: Set StkHld8Db = Db(StkHld8DtaFb): End Function
Function StkHld8TmpPdb() As Database: Set StkHld8TmpPdb = Db(StkHld8TmpFba): End Function
Function StkHld8Frm(F$) As Access.Form: Set StkHld8Frm = StkHld8Acs.Forms(F): End Function
Sub BrwStkHld8Qd(): BrwQd StkHld8Db: End Sub

