Attribute VB_Name = "MxIdeMthLines"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthLines."
#If Doc Then
'
'
#End If

'--Mth
Function MthzMN(M As CodeModule, Mthn, Optional ShtTy$) As String(): MthzMN = MthzN(Src(M), Mthn, ShtTy): End Function
Function MthzN(Src$(), Mthn, Optional ShtTy$) As String():            MthzN = Mth(Src, Mthix(Src, Mthn, ShtTy)): End Function
Function Mth(Src$(), Mthix) As String():                      Mth = AwBE(Src, Mthix, MthEix(Src, Mthix)): End Function
Function Mthl$(Src$(), Mthix):                               Mthl = JnCrLf(Mth(Src, Mthix)):              End Function
Function MthlzN$(Src$(), Mthn, Optional ShtTy$):           MthlzN = Mthl(Src, Mthix(Src, Mthn, ShtTy)):   End Function ' return First Mthl:
Function MthlzMN(M As CodeModule, Mthn, Optional ShtTy$): MthlzMN = MthlzN(Src(M), Mthn, ShtTy):          End Function
Function CMthl$():                                          CMthl = MthlzMN(CMd, CMthn):                  End Function

Function MthlzPN$(P As VBProject, Mthn, Optional ShtMthTy$)
MthlzPN = MthlzMN(MdzMthn(P, Mthn, ShtMthTy), Mthn, ShtMthTy)
End Function

'--Mth-Ay
Function MthlAy$(Src$(), Mthn, Optional ShtMthTy$):                           MthlAy = Mthl(Src, MthixzN(Src, Mthn, ShtMthTy)): End Function
Function MthlAyzMN(M As CodeModule, Mthn, Optional ShtMthTy$) As String(): MthlAyzMN = MthlAyzN(Src(M), Mthn, ShtMthTy):        End Function
Function MthlAyzN(Src$(), Mthn, Optional ShtMthTy$) As String():            MthlAyzN = MthlAyzMN(CMd, Mthn, ShtMthTy):          End Function
Function MthlAyzK(Src$(), MthKn) As String():                               MthlAyzK = AwBei(Src, MthBei(Src, MthKn)):          End Function

Function EmpTstSubl$(Mthn) ' :Lines #Sub-Mth-Lines#
Dim L1$, L2$
L1 = FmtQQ("Private Sub ?__Tst()", Mthn)
L2 = "End Sub"
EmpTstSubl = L1 & vbCrLf & L2
End Function
