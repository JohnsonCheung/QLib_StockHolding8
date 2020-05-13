Attribute VB_Name = "MxIdeMdSts"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdSts."
Type MthCnt
    NLin As Long
    NPub As Integer
    NPrv As Integer
    NFrd As Integer
End Type

Function MthCntzMthln(Mthlny$(), NLin&) As MthCnt
With MthCntzMthln
    .NLin = NLin
    Dim L: For Each L In Itr(Mthlny)
        Select Case Mdy(L)
        Case "", "Public": .NPub = .NPub + 1
        Case "Private":    .NPrv = .NPrv + 1
        Case "Friend":     .NFrd = .NFrd + 1
        Case Else: Stop
        End Select
    Next
End With
End Function

Sub DmpMthCntzSrc(Src$())
DmpMthCnt MthCnt(Src)
End Sub
Sub DmpMthCnt(A As MthCnt)
Debug.Print MthCntStr(A)
End Sub
Function MthCntStr$(A As MthCnt)
With A
MthCntStr = FmtQQ("[NLin NPub NPrv NFrd](? ? ? ?)", .NLin, .NPub, .NPrv, .NFrd)
End With
End Function
Function MthCnt(Src$()) As MthCnt
MthCnt = MthCntzMthln(Mthlny(Src), Si(Src))
End Function

Function MthCntM() As MthCnt
MthCntM = MthCntzM(CMd)
End Function

Function MthCntzM(M As CodeModule) As MthCnt
MthCntzM = MthCnt(Src(M))
End Function
