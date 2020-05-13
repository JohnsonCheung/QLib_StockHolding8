Attribute VB_Name = "MxIdeSrcDimItm"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Dim.DimItm"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDimItm."
':Adj: #Adjective#
':Pgm: :Cml #Program#
':Cd: :Lines #Code# !Cd generated by pgm
':NCd: :Adj #Non-Cd#
Private Sub DimItmAyP__Tst()
Brw DimItmAyP
End Sub
Function DimItmAyP() As String()
DimItmAyP = DimItmAyzP(CPj)
End Function
Function DimItmAyzP(P As VBProject) As String()
DimItmAyzP = DimItmAyzS(SrczP(P))
End Function

Function DimItmAyzS(Src$()) As String()
':DimItm: :S #Dim-Itm# ! The String aft Kw-Dim bef comma
Dim L: For Each L In Itr(Src)
    PushIAy DimItmAyzS, DimItmAyzL(L)
Next
End Function

Private Sub DimLinAy__Tst()
Brw DimLinAy(SrczP(CPj))
End Sub
Function DimStmtAy(Src$()) As String()
Dim S: For Each S In Itr(DimLinAy(Src)): PushI DimStmtAy, DimStmt(S): Next
End Function
Function DimLinAy(Src$()) As String()
'DimLin: :Ln #Dim-Ln# ! the Fst4Chr must be [Dim ]
Dim L: For Each L In Itr(Src): PushNB DimLinAy, DimLin(L): Next
End Function

Function IsLnDim(L) As Boolean
If Not IsCdLn(L) Then Exit Function
Select Case True
Case HasPfx(LTrim(L), "Dim "), HasSubStr(L, " Dim "): IsLnDim = True
End Select
End Function

Function DimLin$(L)
If IsLnDim(L) Then DimLin = L
End Function

Function DimItmAy(DimStmt$) As String()
If NoPfx(DimStmt, "Dim ") Then Exit Function
DimItmAy = AmTrim(SplitComma(Mid(DimStmt, 5)))
End Function

Function DimStmt$(DimLin)
Const CSub$ = CMod & "DimStmt"
Dim P%: P = InStr(DimLin, "Dim ")
If P = 0 Then Thw CSub, "Given DimLin is not DimLin", "DimLin", DimLin
Dim O$: O = Mid(DimLin, P)
DimStmt = BefOrAll(BefOrAll(O, "'"), ":")
End Function

Function DimItmAyzL(L) As String()
Dim D$: D = DimLin(L): If D = "" Then Exit Function
Dim S$: S = DimStmt(D)
DimItmAyzL = DimItmAy(S)
End Function