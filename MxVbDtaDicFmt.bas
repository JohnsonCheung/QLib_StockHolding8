Attribute VB_Name = "MxVbDtaDicFmt"
Option Compare Text
Option Explicit
Const CNs$ = "Dic"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbDtaDicFmt."

Sub BrwDic__Tst()
Dim R As DAO.Recordset
Set R = Rs(DutyDtaDb, "Select Sku,BchNo from PermitD where BchNo<>''")
BrwDic JnStrDicTwoFldRs(R), True
End Sub

Sub VcDic(A As Dictionary, Optional InlValTy As Boolean, Optional ExlIx As Boolean, Optional FnPfx$ = "Dic_")
BrwDic A, InlValTy, ExlIx, FnPfx
End Sub

Sub BrwDic(A As Dictionary, Optional InlValTy As Boolean, Optional ExlIx As Boolean, Optional FF$ = "Key Val", Optional FnPfx$ = "Dic_")
BrwAy FmtDic(A, InlValTy, FF), FnPfx
End Sub

Private Sub DmpDic__Tst()
DmpDic ZZSamp1, True
End Sub

Sub DmpDic(A As Dictionary, Optional InlValTy As Boolean, Optional FF$ = "Key Val", Optional BegIx% = 1, Optional Fmt As eTblFmt, Optional Tit$)
D FmtDic(A, InlValTy, FF, BegIx, Fmt, Tit)
End Sub

Function FmtDic(D As Dictionary, Optional InlValTy As Boolean, Optional H12$ = "Key Val", Optional BegIx% = 1, Optional Fmt As eTblFmt, Optional Tit$) As String()
FmtDic = FmtS12y(S12yzDic(ZfmtAddValTy(D, InlValTy)), H12, BegIx, Fmt, Tit)
End Function

Private Function ZfmtAddValTy(D As Dictionary, InlValTy As Boolean) As Dictionary ' ret nwDic with optionally K added with VarTy
If Not InlValTy Then Set ZfmtAddValTy = D: Exit Function
Dim K$(): K = ZfmtNwKy(D)
Dim V$(): V = SyzDii(D)
Set ZfmtAddValTy = DiczAy2(K, V)
End Function

Private Function ZfmtNwKy(D As Dictionary) As String() 'ret SKy with VarTy of Dii added
Dim K$(): K = AmAli(SyzDik(D))
Dim T$(): T = TynyzDic(D)
ZfmtNwKy = AddAliStrCol(K, T)
End Function

Private Function ZZSamp1() As Dictionary
Set ZZSamp1 = New Dictionary
ZZSamp1.Add "A", 1
ZZSamp1.Add "B", 2
ZZSamp1.Add "C", 3&
End Function
