Attribute VB_Name = "MxIdeSrcStmtInsp"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Stmt"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcStmtInsp."

Sub Insp(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
Dim F$: If Fun <> "" Then F = " (@" & Fun & ")"
Dim A$(): A = BoxzS("Insp: " & Msg & F)
BrwAy Sy(A, FmtNav(Nav))
End Sub
