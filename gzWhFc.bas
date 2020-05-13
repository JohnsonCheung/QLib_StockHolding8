Attribute VB_Name = "gzWhFc"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzWhFc."

Function AndStm$(Stm$): AndStm = QpAndFeq("Stm", Stm): End Function
Function AndCo(Co As Byte): AndCo = QpAndFeq("Co", Co): End Function

Function WhereFcStm$(A As StmYM)
WhereFcStm = WhereFc(A.Y, A.M) & AndStm(A.Stm)
End Function

Function WhereFcCoStm$(Co As Byte, A As StmYM)
WhereFcCoStm = WhereFcStm(A) & AndCo(Co)
End Function

Function WhereFc$(Y As Byte, M As Byte)
WhereFc = Wh("VerYY=" & Y & " and VerMM=" & M)
End Function

Function WhereFcCo$(A As CoYM)
WhereFcCo = WhereFc(A.Y, A.M) & AndCo(A.Co)
End Function
