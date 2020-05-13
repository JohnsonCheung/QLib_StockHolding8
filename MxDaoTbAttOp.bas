Attribute VB_Name = "MxDaoTbAttOp"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CNs$ = "Att"
Const CMod$ = CLib & "MxDaoTbAttOp."
'**CrtTbAtt
Sub CrtTbAtt(D As Database): CrtTbAttNrm D: End Sub

'**CrtTbAtt-Nrm
Sub CrtCTbAttNrm__Tst()
DrpCTbAtt
CrtCTbAttNrm
End Sub
Sub CrtCTbAttNrm(): CrtTbAttNrm CDb: End Sub
Sub CrtTbAttNrm(D As Database)
D.TableDefs.Append NwAttTd
D.TableDefs.Append NwAttdTd
End Sub
Private Function NwAttTd() As DAO.TableDef
Dim O As New DAO.TableDef
With O
    .Name = "Att"
    .Fields.Append FdzId("AttId")
    .Fields.Append FdzNNTxt("Attn")
    .Fields.Append FdzAtt("Att")
    .Indexes.Append NwSkIdxzF(O, "Attn")
End With
Set NwAttTd = O
End Function
Private Function NwAttdTd() As DAO.TableDef
Dim O As New DAO.TableDef
With O
    .Name = "Attd"
    .Fields.Append FdzId("AttdId")
    .Fields.Append FdzNNLng("AttId")
    .Fields.Append FdzNNTxt("Fn")
    .Fields.Append FdzNNDte("FilTim")
    .Fields.Append FdzNNLng("FilSi")
    .Indexes.Append NwSkIdxzFF(O, "AttId Fn")
End With
Set NwAttdTd = O
End Function

'**CrtTbAttSchm
Sub CrtTbAttSchm(D As Database): CrtSchm D, AttSchm: End Sub
Function AttSchm() As String()
Const A$ = "Tbl"
Const B1$ = "  Att * *n | Att"
Const B2$ = "  Attd * AttId Fn | FIlTim FilSi"
Const C$ = "EleFld"
Const D$ = "  T22 FilTimSi22"
Const E$ = "  Att AttFn"
Const F$ = "  Nm  Attn"
AttSchm = SyzAp(A, B1, B2, C, D, E, F)
End Function

'**EnsTbAtt
Private Sub EnsTbAtt__Tst()
EnsTbAtt CDb
BrwDb CDb
End Sub
Sub EnsCTbAtt(): EnsTbAtt CDb: End Sub
Sub EnsTbAtt(D As Database): EnsSchm D, AttSchm: End Sub

'**DrpTbAtt
Sub DrpTbAtt(D As Database): DrpTT D, "Att Attd": End Sub
Sub DrpCTbAtt(): DrpTbAtt CDb: End Sub

