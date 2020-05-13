Attribute VB_Name = "MxVbStrLinesInf"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str.CntSi"
Const CMod$ = CLib & "MxVbStrLinesInf."
Function LinesInf$(Lines$)
LinesInf = FmtQQ("Cnt-Si(?-?)", LnCnt(Lines), Len(Lines))
End Function

Function LinesInfzLy$(Ly$())
LinesInfzLy = LinesInf(JnCrLf(Ly))
End Function
