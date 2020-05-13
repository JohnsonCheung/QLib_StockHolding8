Attribute VB_Name = "gzWhOH"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzWhOH."
Private Const OH3YmdBexpTp$ = "{A}YY={Y} and {A}MM={M} and {A}DD={D}"
Function OH3YmdBexp$(Y As Byte, M As Byte, D As Byte, Optional Alias0$): OH3YmdBexp = FmtMacro(OH3YmdBexpTp, "A Y M D", Alias(Alias0), Y, M, D): End Function
Function OHDteBexp$(D As Date, Optional Alias$): OHDteBexp = OHYmdBexp(YmdzDte(D), Alias): End Function
Function CoOHYmdBexp$(A As CoYmd): CoOHYmdBexp = OHYmdBexp(A.Ymd) & " and Co=" & A.Co: End Function
Function OHYmdBexp$(A As Ymd, Optional Alias$)
With A
OHYmdBexp = OH3YmdBexp(.Y, .M, .D, Alias)
End With
End Function
Function StmBexp$(Stm$):        StmBexp = WhFeq("Stm", Stm): End Function
Function LasOHBexp$():        LasOHBexp = OHYmdBexp(LasOHYmd): End Function
Function LasOHyymmdd&():    LasOHyymmdd = VzCQ("Select Max(YY*10000+MM*100+DD) from OH"): End Function
Function LasOHDte() As Date:   LasOHDte = DtezYmd(LasOHYmd):                              End Function
Function LasOHYmd() As Ymd:    LasOHYmd = YmdzYYMMDD(LasOHyymmdd):                        End Function

Function WhLasOH$(): WhLasOH = OHYmdBexp(LasOHYmd): End Function
Function WhOHYmd$(A As Ymd, Optional Alias$): WhOHYmd = " where " & OHYmdBexp(A, Alias): End Function

