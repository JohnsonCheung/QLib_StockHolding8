Attribute VB_Name = "MxVbAyFmt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbAyFmt."
Sub DmpCntAy__Tst()
DmpCntAy MthnyP
End Sub
Sub DmpCntAy(Ay, Optional C As VbCompareMethod = VbCompareMethod.vbTextCompare)
Dmp FmtCntDi(CntDi(Ay, C))
End Sub
