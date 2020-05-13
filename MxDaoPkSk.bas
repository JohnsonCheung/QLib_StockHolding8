Attribute VB_Name = "MxDaoPkSk"
Option Compare Text
Option Explicit
Const CNs$ = "Def"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoPkSk."

Sub ChkPk(D As Database, T)
ChkEr Sy(PkEmsgl(D, T))
End Sub

Function PkEmsgl$(D As Database, T)
If HasPk(D, T) Then Exit Function
Dim Pk$(): Pk = PkFny(D, T)
Select Case True
Case Si(Pk) = 0: PkEmsgl = FmtQQ("[?] does not have PrimaryKey-Idx", T)
Case Si(Pk) <> 1: PkEmsgl = FmtQQ("There is PrimaryKey-Idx, but it has [?] fields[?]", Si(Pk), Tml(Pk))
Case Pk(0) <> T & "Id": PkEmsgl = FmtQQ("There is One-field-PrimaryKey-Idx of Fldn(?), but it should named as ?Id", Pk(0), T)
Case FdzF(D, T, 0).Name <> T & "Id": PkEmsgl = FmtQQ("The Pk-field(?Id) should be first fields, but now it is (?)", T, FdzF(D, T, T & "Id").OrdinalPosition)
End Select
End Function
Function PkSkEmsgzAllTbl(D As Database) As String()
Dim T$, I
For Each I In Tny(D)
    T = I
    PushIAy PkSkEmsgzAllTbl, PkSkEmsg(D, T)
Next
End Function

Function PkSkEmsg(D As Database, T) As String()
PushS PkSkEmsg, PkEmsgl(D, T)
ChkSk D, T
End Function

Function ChkSk$(D As Database, T)
If Not HasSk(D, T) Then
    ChkSk = FmtQQ("Not SecondaryKey for Table[?] in Db[?]", T, D.Name)
    Exit Function
End If
Dim Sk As DAO.Index, I As DAO.Index
Set Sk = SkIdx(D, T)
Select Case True
Case Not Sk.Unique
    ChkSk = FmtQQ("SecondaryKey is not unique for Table[?] in Db[?]", T, D.Name)
Case Else
    Set I = FstUniqIdx(D, T)
    If Not IsNothing(I) Then
 '       ChkSk = FmtQQ("No SecondaryKey, but there is uniq idx, it should name as SecondaryKey for Table[?] Db[?] UniqIdxn[?] IdxFny[?]", _
            T, D.Name, I.Name, JnTermy(FnyzIdx(I)))
    End If
End Select
End Function

Function ChkSsk$(D As Database, T)
Dim O$, Sk$(): Sk = SkFny(D, T)
O = ChkSk(D, T): If O <> "" Then ChkSsk = O: Exit Function
If Si(Sk) <> 1 Then
'    ChkSsk = FmtQQ("Secondary is not single field. Tbl[?] Db[?] SkFfn[?]", T, D.Name, JnTermy(Sk))
End If
End Function
