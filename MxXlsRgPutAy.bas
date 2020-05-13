Attribute VB_Name = "MxXlsRgPutAy"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CNs$ = "Xls.Put.Ay"
Const CMod$ = CLib & "MxXlsRgPutAy."
Sub PutAyh(Ayh, At As Range)
NwLozSq Sqh(Ayh), At
End Sub

Sub PutAyv(Ayv, At As Range)
NwLozSq Sqv(Ayv), At
End Sub
Sub PutHss(Hss$, At As Range)
PutAyh SyzSS(Hss), At
End Sub
Sub PutVss(Vss$, At As Range)
PutAyv SyzSS(Vss), At
End Sub

Sub PutSSV(SSV$, At As Range)
PutAyv SyzSS(SSV), At
End Sub
