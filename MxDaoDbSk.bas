Attribute VB_Name = "MxDaoDbSk"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbSk."
Const CNs$ = "Dao.Ssk"
Public Const Pkn$ = "PrimaryKey"

Sub DltReczNotInSskVet(D As Database, SskTbn$, NotInSskVet As Dictionary) _
'Delete Db-T record for those record's Sk not in NotInSSskv, _
'Assume T has single-fld-sk
Const CSub$ = CMod & "DltRecNotInSskv"
Dim Sskf$: Sskf = SskFldn(D, SskTbn)
If NotInSskVet.Count = 0 Then Thw CSub, "Given NotInSskVet cannot be empty", "Db SskTbn SskFldn", D.Name, SskTbn, Sskf
Dim Q$, ExcessVet As Dictionary
Set ExcessVet = MinusAet(SskVet(D, SskTbn), NotInSskVet)
If ExcessVet.Count = 0 Then Exit Sub
RunSqy D, SqyDlt_T_WhFld_InAet(SskTbn, Sskf, ExcessVet)
End Sub

Sub InsReczSskVet(D As Database, SskTbn, ToInsSskVet As Dictionary) _
'Insert Single-Field-Secondary-Key-Aet into Dbt
'Assume T has single-fld-sk and can be inserted by just giving such SSk-value
Dim ShouldInsVet As Dictionary
    Set ShouldInsVet = MinusAet(ToInsSskVet, SskVet(D, SskTbn))
If ShouldInsVet.Count = 0 Then Exit Sub
Dim F$: F = SskFldn(D, SskTbn)
With RszT(D, SskTbn)
    Dim I: For Each I In ShouldInsVet
        .AddNew
        .Fields(F).Value = I
        .Update
    Next
    .Close
End With
End Sub

Function CSkFny(T) As String()
CSkFny = SkFny(CDb, T)
End Function

Function SkFny(D As Database, T) As String()
SkFny = Itn(SkIdx(D, T).Fields)
End Function

Function SkFnyzTd(T As DAO.TableDef) As String()
SkFnyzTd = Itn(T.Indexes(T.Name).Fields)
End Function

Function SkIdx(D As Database, T) As DAO.Index
Set SkIdx = Idx(D, T, T)
End Function

Function NwIdx(Td As DAO.TableDef, K$, Fny$(), Optional IsPk As Boolean, Optional IsUKy As Boolean) As DAO.Index
Dim O As DAO.Index: Set O = Td.CreateIndex(K)
If IsPk Then
    O.Primary = True
    O.Unique = True
ElseIf IsUKy Then
    O.Unique = True
End If
Dim Fds As DAO.IndexFields: Set Fds = O.Fields
Dim F: For Each F In Fny
    Fds.Append O.CreateField(F)
Next
Set NwIdx = O
End Function

Function NwSkIdxzF(Td As DAO.TableDef, F) As DAO.Index
Set NwSkIdxzF = NwSkIdx(Td, Sy(F))
End Function

Function NwSkIdxzFF(Td As DAO.TableDef, FF$) As DAO.Index
Set NwSkIdxzFF = NwSkIdx(Td, FnyzFF(FF))
End Function

Function NwSkIdx(Td As DAO.TableDef, Fny$()) As DAO.Index
Set NwSkIdx = NwIdx(Td, Td.Name, Fny, IsUKy:=True)
End Function

Function SskFldn$(D As Database, T)
Const CSub$ = CMod & "SskFld"
Dim Sk$(): Sk = SkFny(D, T): If Si(Sk) = 1 Then SskFldn = Sk(0): Exit Function
Thw CSub, "SkFny-Sz<>1", "Db T, SkFny-Si SkFny", D.Name, T, Si(Sk), Sk
End Function

Function SskVet(D As Database, T) As Dictionary
'SskVet is [S]ingleFielded [S]econdKey [K]ey [V]alue S[et], which is always a Value-Aet.
'and Ssk is a field-name from , which assume there is a Unique-Index with name "SecordaryKey" which is unique and and have only one field
Set SskVet = AetzF(D, T, SskFldn(D, T))
End Function
