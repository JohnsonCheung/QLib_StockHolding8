Attribute VB_Name = "MxAcsOp"
Option Explicit
Option Compare Text
Const CNs$ = "Acs.Op"
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsOp."

Private Sub OpnFb__Tst()
Dim A As Database: Set A = NwAcs.CurrentDb
Stop
End Sub

Sub OpnFrm(FrmNm$)
DoCmd.OpenForm FrmNm
End Sub

Sub OpnFb(A As Access.Application, Fb)
If FbzAcs(A) = Fb Then Exit Sub
ClsAcsDb A
A.OpenCurrentDatabase Fb
End Sub

Sub ClsAcsDb(A As Access.Application)
On Error Resume Next
A.CurrentDb.Close
End Sub

Sub BrwFb(Fb)
Static Acs As New Access.Application
OpnFb Acs, Fb
Acs.Visible = True
End Sub

Sub BrwCQ(Q$, Optional QryNmPfx$ = "Q")
BrwQ CDb, Q, QryNmPfx
End Sub

Sub BrwQ(D As Database, Q$, Optional QryNmPfx$ = "Q")
AcszDb(D).DoCmd.OpenQuery TmpQry(D, Q, QryNmPfx)
End Sub

Sub BrwCT(T)
BrwT CDb, T
End Sub

Sub BrwT(D As Database, T)
Dim A As Access.Application: Set A = AcszFb(D.Name)
A.DoCmd.OpenTable T
End Sub

Sub CBrwTT(TT$)
BrwTT CDb, TT
End Sub

Sub BrwTT(D As Database, TT$)
Dim T: For Each T In ItrzTml(TT)
    BrwT D, T
Next
End Sub

Sub ClsAllCTbl()
ClsAllTbl Acs
End Sub

Sub ClsAllTbl(A As Access.Application)
Dim T: For Each T In Itr(TnyzA(A))
    A.DoCmd.Close acTable, T
Next
End Sub

Sub ClsAcsTbl(A As Access.Application, T)
A.DoCmd.Close acTable, T, acSaveYes
End Sub
Sub ClsDbt(D As Database, T)
ClsAcsTbl AcszDb(D), T
End Sub

Sub ClsDbtTT(D As Database, TT$)
Dim A As Access.Application: Set A = AcszDb(D)
Dim T: For Each T In Termy(TT)
    ClsAcsTbl A, T
Next
End Sub

Sub ClsTzA(A As Access.Application, T)
A.DoCmd.Close acTable, T
End Sub

Sub ClsAllFrm(A As Access.Application)
While A.Forms.Count > 0
A.DoCmd.Close acForm, A.Forms(0).Name, acSaveNo
Wend
End Sub

Sub ClrCAcsSts()
ClrAcsSts Acs
End Sub

Sub ClrAcsSts(A As Access.Application)
A.SysCmd acSysCmdClearStatus
End Sub

Sub OpnTblRO(T): DoCmd.OpenTable T, acViewNormal, acReadOnly: End Sub
Sub OpnTbl(T): DoCmd.OpenTable T: End Sub
