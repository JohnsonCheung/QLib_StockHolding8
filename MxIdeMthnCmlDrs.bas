Attribute VB_Name = "MxIdeMthnCmlDrs"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthnCmlDrs."
':MthCml$ = "NewType:Sy."

Function MthCmlDrsP() As Drs
Dim A$()
A = MthnyP
A = AeSfx(A, "__Tst")
A = AePfx(A, "T_")
A = AwDis(A)
A = QSrt(A)
MthCmlDrsP = CmlnnDrs(A)
End Function

Function MthCmlWsP() As Worksheet: Set MthCmlWsP = WszDrs(MthCmlDrsP): End Function
