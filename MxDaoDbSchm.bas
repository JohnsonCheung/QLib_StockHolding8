Attribute VB_Name = "MxDaoDbSchm"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbSchm."
'**EnsSchm
Private Sub EnsCSchm__Tst()
Dim Schm$()
YY:
    Schm = SampSchm(1)
    GoTo Tst
Tst:
    EnsCSchm Schm
    Stop
    Return
End Sub

Sub EnsCSchm(Schm$()): EnsSchm CDb, Schm: End Sub
Sub EnsSchm(D As Database, Schm$())

End Sub

'**CrtSchm
Sub CrtSchm(D As Database, Schm$())
Const CSub$ = CMod & "CrtSchm"
Dim A As SchmSrc: A = SchmSrczS(Schm)
Dim B As SchmPsr: B = SchmPsrzS(A)
Dim Er$(): Er = ErzSchmEr(B.Er)
ChkEr Er, CSub
With SchmBldzD(B.Dta)
    Stop
    AppTdAy D, .TdAy
    RunSqy D, .PkSqy
    RunSqy D, .SkSqy
    RunSqy D, .KeySqy
    RunSqy D, .FkSqy
    SetTblDeszDi D, .TblDesDi
    SetFldDeszDi D, .FldDesDi
End With
End Sub
Sub CrtCSchm(Schm$()): CrtSchm CDb, Schm: End Sub
