Attribute VB_Name = "MxXlsLoLofl"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CNs$ = "Lof"
Const CMod$ = CLib & "MxXlsLoLofl."
':Lofl: :Lines #Lo-Fmtr-Lines#
Function LoflzQt$(A As Excel.QueryTable)
LoflzQt = LoflzFbtStr(FbtStrzQt(A))
End Function

Function LoflzT$(D As Database, T)
LoflzT = TblPrp(D, T, "Lofl")
End Function

Sub SetLoflzT(D As Database, T, V$)
SetTbPrp D, T, "Lofl", V
End Sub

Function LoflzLo$(A As ListObject)
LoflzLo = LoflzQt(LoQt(A))
End Function

Function LoflzFbt$(Fb, T)
LoflzFbt = LoflzT(Db(Fb), T)
End Function

Sub SetLoflzFbt(Fb, T, LoflzVbl$)
SetLoflzT Db(Fb), T, LoflzVbl
End Sub

Function LoflzFbtStr$(FbtStr$)
Dim Fb$, T$
AsgFbtStr FbtStr, Fb, T
LoflzFbtStr = LoflzFbt(Fb, T)
End Function
