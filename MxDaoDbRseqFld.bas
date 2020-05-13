Attribute VB_Name = "MxDaoDbRseqFld"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoDbRseqFld."

Sub NewRseqTbl(T, Nseq$(), OrdBy$)
Dim F$: F = Join(AmQuoSq(Nseq), ",")
Dim SelF$: SelF = "Select " & F
Dim Into$: Into = " Into [@" & T & "]"
Dim Fm$: Fm = " From [#" & T & "]"
Dim Sql$: Sql = SelF & Into & Fm & " Order By " & OrdBy
RunCQ Sql
End Sub

Sub SrtCFld(T, FF$)
SrtFld CDb, T, FF
End Sub
Function IsLnkTbl(D As Database, T) As Boolean
IsLnkTbl = Td(D, T).Connect <> ""
End Function
Sub SrtFld(D As Database, T, FF$)
Const CSub$ = CMod & "RseqFld"
If IsLnkTbl(D, T) Then Thw CSub, "Given table is a linked, cannot Rseq", "T", T
'#1 Chk ExcF
'#2 Mk  NewF
'#3 Rseq
'#4 Chk if sequenced
Dim GivF$(): GivF = FnyzFF(FF)
Dim OldF$(): OldF = Fny(D, T)
Dim ExcF$(): ExcF = MinusAy(GivF, OldF)
If Si(ExcF) > 0 Then Thw CSub, "Given FF some excess field than given T", "Excess-field Given-FF T-FF T", Tml(ExcF), FF, Tml(OldF), T

Dim MisF$(): MisF = MinusAy(OldF, GivF)
Dim NewF$(): NewF = AddSy(GivF, MisF)

Dim Td1 As DAO.TableDef: Set Td1 = Td(D, T)
Dim BefF$(): BefF = Itn(Td1.Fields)
Dim M%: M = MaxOrdinalPosition(D, T)
Dim J%: For J = UB(NewF) To 0 Step -1
    Td1.Fields(NewF(J)).OrdinalPosition = M + 1 + J
Next
J = 0
Dim F As DAO.Field: For Each F In Td1.Fields
    J = J + 1
    F.OrdinalPosition = J
Next

For J = 0 To UB(NewF)
    If Td1.Fields(J).OrdinalPosition <> J + 1 Then
        Thw CSub, "Table not reseq as expected", "Given-FF Bef-Srt-Given-Tbl-FF Aft-Srt-Given-Tbl-FF Aft-Srt-Given-Tbl-OrdinalPosition", FF, Tml(BefF), Tml(Itn(Td1.Fields)), FmtOrdinalPosition(D, T)
    End If
Next
End Sub

Function CMaxOrdinalPosition%(T): CMaxOrdinalPosition = MaxOrdinalPosition(CDb, T): End Function

Function MaxOrdinalPosition%(D As Database, T)
MaxOrdinalPosition = MaxItp(Td(D, T).Fields, "OrdinalPosition")
Exit Function
Dim Tbl As DAO.TableDef: Set Tbl = Td(D, T)
Dim O&, M%
Dim F As DAO.Field: For Each F In Tbl.Fields
    M = F.OrdinalPosition
    If M > O Then O = M
Next
MaxOrdinalPosition = O
End Function
