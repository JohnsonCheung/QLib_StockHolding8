Attribute VB_Name = "MxVbStrReadBql"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbStrReadBql."
Const CNs$ = "Dao.Bql"
Const ShtTyBql$ = "Short-Type-Si-Colon-Fldn-Bql:Sht.Ty.s.c.f.Bql: It is a [Bql] with each field is a [ShtTyscf]"
':Bql: :Ln #Back-Quo-Line# ! Back-Quo is (`) and it is a String.  Each field is separated by (`)
':Fbql: :Ft #Fullfilename-Bql# ! Each line is a [Bql]|Fst line is [ShtTyBql]
':ShtTys: :Nm #ShtTy-Size# ! It is a [ShtTy] or (Tnnn) where nnn can 1 to 3 digits of value 1-255"
':ShtTyLis: :Cml #Short-Type-List# ! Each :Cml is 1 or 3 chr of :ShtTy
':ShtTyscf: :Term #ShtTy-Si-Colon-Fldn#  ! If Fldn have space, then ShtTyscf should be sq bracket"
':ShtTyBql: :Bql #ShtTyscf-Bql# ! Each field is a [ShtTyscf].  It is used to create an empty table by CrtTblzShtTyscfBql"

Function ShtTyscfBqlzDrs$(A As Drs)
Dim Dy(): Dy = A.Dy
If Si(Dy) = 0 Then ShtTyscfBqlzDrs = Jn(A.Fny, "`"): Exit Function
Dim O$(), F$, I, C&, Fny$()
Fny = A.Fny
For C = 0 To NColzDrs(A) - 1
    F = Fny(C)
    PushI O, ShtTyscfzCol(ColzDy(Dy, C), F)
Next
ShtTyscfBqlzDrs = Jn(O, "`")
End Function

Function ShtTyscfzCol$(Col(), F$)
Dim O$: O = AddNBAp(ShtTyszCol(Col), ":") & F
If IsNeedQuo(F) Then O = QuoSq(O)
ShtTyscfzCol = O
End Function

Private Sub CrtTTzBqlPth__Tst()
Dim D As Database: Set D = TmpDb
Dim P$: P = TmpPthi
WrtFbqlzDb P, DutyDtaDb
CrtTTzBqlPth D, P
BrwDb D
End Sub

Sub CrtTTzBqlPth(D As Database, BqlPth$)
CrtTTzBqlPthFnny D, BqlPth, FnnAy(BqlPth, "*.bql.txt")
End Sub

Sub CrtTTzBqlPthFnny(D As Database, BqlPth$, FnnAy$())
Dim T, P$, Fbql$
P = EnsPthSfx(BqlPth)
For Each T In FnnAy
    Fbql = P & T & ".txt"
    CrtTblzFbql D, Fbql
Next
End Sub

Private Sub CrtTblzFbql__Tst()
Dim Fbql$: Fbql = TmpFt
WrtFbql Fbql, DutyDtaDb, "PermitD"
Dim D As Database: Set D = TmpDb
CrtTblzFbql D, "PermitD", Fbql
BrwDb D
Stop
End Sub

Sub CrtFbzBqlPth(BqlPth$, Optional Fb0$)
Dim Fb$
    Fb = Fb0
    If Fb = "" Then Fb = BqlPth & Fdr(BqlPth) & ".accdb"
DltFfnIf Fb
CrtFb Fb
Dim D As Database, IFfn, T$
Set D = Db(Fb)
For Each IFfn In Ffny(BqlPth, "*.bql.txt")
    CrtTblzFbql D, IFfn
Next
End Sub

Function TzFbql$(Fbql)
Const CSub$ = CMod & "TzFbql"
If Not HasSfx(Fbql, ".bql.txt") Then Thw CSub, "Fbql does not have .bql.txt sfx", "Fbql", Fbql
TzFbql = RmvSfx(Fn(Fbql), ".bql.txt")
End Function

Sub CrtTblzFbql(D As Database, Fbql, Optional T0$)
Dim T$
    T = T0
    If T = "" Then T = TzFbql(Fbql)

Dim F%, L$, R As DAO.Recordset
F = FnoI(Fbql)
Line Input #F, L
CrtTblzShtTyscfBql D, T, L

Set R = RszT(D, T)
While Not EOF(F)
    Line Input #F, L
    InsRszBql R, L
Wend
R.Close
Close #F
End Sub

Sub CrtTblzShtTyscfBql(D As Database, T, ShtTyscfBql$)
Dim Td As New DAO.TableDef
Td.Name = T
Dim I
For Each I In Split(ShtTyscfBql, "`")
    Td.Fields.Append FdzShtTyscf(I)
Next
D.TableDefs.Append Td
End Sub

Function FdzShtTyscf(ShtTyscf) As DAO.Field
Dim T As DAO.DataTypeEnum
Dim S As Byte
With Brk2(ShtTyscf, ":")
    Select Case True
    Case .S1 = "":                 T = dbText: S = 255
    Case FstChr(ShtTyscf) = "T":   T = dbText: S = RmvFstChr(.S1)
    Case Else:                     T = DaoTyzShtTy(.S1)
    End Select
    Dim ZLen As Boolean: ZLen = T = dbText
    Set FdzShtTyscf = Fd(.S2, T, TxtSi:=S, ZLen:=ZLen)
End With
End Function

Function ShtTyBqlzT$(D As Database, T)
Dim Ay$(), F As DAO.Field
For Each F In D.TableDefs(T).Fields
    PushI Ay, ShtTyszFd(F) & ":" & F.Name
Next
ShtTyBqlzT = Jn(Ay, "`")
End Function

Function ShtTyszFd$(A As DAO.Field)
Dim B$: B = ShtDaoTy(A.Type)
If A.Type = dbText Then
    B = B & A.Size
End If
ShtTyszFd = B
End Function
