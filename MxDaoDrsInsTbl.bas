Attribute VB_Name = "MxDaoDrsInsTbl"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDrsInsTbl."

Sub InsTblzDrs(D As Database, T, B As Drs)
Dim F$(): F = IntersectAy(Fny(D, T), B.Fny)
InsRszDy RszTFny(D, T, F), DrszSelFny(B, F).Dy
End Sub

Sub InsTblzDy(D As Database, T, Dy())
InsRszDy RszT(D, T), Dy
End Sub
