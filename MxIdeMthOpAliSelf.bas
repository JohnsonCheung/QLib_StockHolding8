Attribute VB_Name = "MxIdeMthOpAliSelf"
Option Compare Text
Option Explicit
Const CNs$ = "Mth.Ali.Self"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthOpAliSelf."

Sub AliMthzSelf()
'Cpy Md
    Const TMdn$ = "QIde_B_AliMth"      ' #The-Mdn
    Const TmMdn$ = "ATmp"                ' #Tmp-Mdn
:                                    EnsCls CPj, TmMdn
    Dim FmM As CodeModule: Set FmM = Md(TMdn)
    Dim ToM As CodeModule: Set ToM = Md(TmMdn)
    Dim OIsCpy As Boolean:  OIsCpy = CpyMd(FmM, ToM)
:                                    If OIsCpy Then MsgBox "Copied": Exit Sub

'Ali
    Const TMthn$ = "AliMthByLno"      ' #The-Mthn
    Dim M As CodeModule: Set M = Md(TMdn)
    Dim Lno&:        Lno = Mthlno(M, TMthn)
    'ATmp.AliMthByLno M, Mthlno, Upd:=eUpdAndRpt, IsUpdSelf:=True
End Sub

Sub AliMthlnkEr()
Dim M As CodeModule: Set M = Md("QDao_Lnk_LnkEr")
Dim L&:                  L = Mthlno(M, "LnkEr")
:                            AliMthzLno M, L
End Sub
