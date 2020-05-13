Attribute VB_Name = "gzMB52Action"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzMB52Action."
Private Const MB52ITbl$ = ">MB52"
Private Const GitITbl$ = ">Git"
Private Const sqlUpdTbReportToNoGit$ = "Update Report Set GitLoadDte=Null,GitNRec=Null,GitBtl=Null,GitSc=Null,GitAc=Null,GitHKD=Null"
Private Type OHTot: NRec As Integer: TotBtl As Long: TotVal As Currency: End Type


'-- Lnk =================
Sub LnkMB52(A As Ymd): CLnkFxw MB52IFx(A), "Sheet1", MB52ITbl: End Sub
Sub LnkGit(A As Ymd):  CLnkFxw GitIFx(A), "Sheet1", GitITbl:   End Sub

'-- LnkAndOpn ===========
Sub LnkAndOpnGitIFx(A As Ymd):  LnkGit A:  OpnTblRO GitITbl:  End Sub
Sub LnkAndOpnMB52IFx(A As Ymd): LnkMB52 A: OpnTblRO MB52ITbl: End Sub

'-- CpyShTo =====================
Private Sub CpySHldTo__Tst()
CpySHldTo LasOHYmd
End Sub

Sub CpySHldTo(A As Ymd)
Dim O86$: O86 = ShOFx(86, A): ChkFfnExist O86
Dim O87$: O87 = ShOFx(87, A): ChkFfnExist O87
XChkCpy12
If IsCpy1Pm Then
    CpyFfnIfDif O86, ShOFx1(86, A)
    CpyFfnIfDif O87, ShOFx1(87, A)
End If
If IsCpy2Pm Then
    CpyFfnIfDif O86, ShOFx2(86, A)
    CpyFfnIfDif O87, ShOFx2(87, A)
End If
End Sub

'-- CpMB52To =====================
Sub CpyMB52To__Tst()
CpyMB52To LasOHYmd
End Sub

Sub CpyMB52To(A As Ymd):
ChkFfnExist MB52OFx(A)
XChkCpy12
Dim OFx$: OFx = MB52OFx(A)
If IsCpy1Pm Then
    CpyFfnIfDif OFx, MB52OFx1(A)
End If
If IsCpy2Pm Then
    CpyFfnIfDif OFx, MB52OFx2(A)
End If
End Sub

'-- SetCpyToPth 1
Sub VisCpy12()
XVisCpy1
XVisCpy2
End Sub

'-- Sel Inp Pth
Sub SelMB52InpPth():    SelCPthPm "MB52_InpPth", XFrm.xInpPth:    End Sub

'-- Sel Pth CpyToPth ============
Sub Sel_CpyToPth1(): SelCPthPm "MB52_CpyToPth1", XFrm.xCpyToPth1: End Sub
Sub Sel_CpyToPth2(): SelCPthPm "MB52_CpyToPth2", XFrm.xCpyToPth2: End Sub

'-- Tgl IsCpyToPth
Sub TglIsCpy1Pm()
TglCPm "MB52_IsCpyToPth1"
XVisCpy1
End Sub

Sub TglIsCpy2Pm()
TglCPm "MB52_IsCpyToPth2"
XVisCpy2
End Sub

Function IsMB52Loaded(A As Ymd) As Boolean
With CurrentDb.OpenRecordset("Select Top 1 Count(*) from OH" & OHYmdBexp(A))
    If Nz(.Fields(0).Value, 0) = 0 Then .Close: MsgBox "Please [Load MB52] first": Exit Function
    .Close
End With
IsMB52Loaded = True
End Function

'-- ShwDbOHTot
Private Sub ShwDbOHTot__Tst(): ShwDbOHTot SampDDec24: End Sub
Sub ShwDbOHTot(A As Ymd): W1Shw A, W1DbOH(A): End Sub

Private Function W1Sql$(A As Ymd): W1Sql = "Select Count(*) as NRec,Sum(Blt) as TotOH,Sum(V) as TotVal" & OHYmdBexp(A): End Function
Private Function W1DbOH(A As Ymd) As OHTot: W1DbOH = W1DbOHzRs(CRs(W1Sql(A))): End Function
Private Sub W1Shw(A As Ymd, B As OHTot): MsgBox W1Msg(A, B): End Sub
Private Function W1DbOHzRs(Rs As DAO.Recordset) As OHTot
Dim O As OHTot:
With Rs:
    O.NRec = !NRec:
    O.TotBtl = !TotBtl:
    O.TotVal = !TotVal:
End With:
W1DbOHzRs = O
End Function
Private Function W1Msg$(A As Ymd, B As OHTot)
With A
W1Msg = FmtStr("Date[20{0}-{1}-{2}]" & vbCrLf & "NRec[{3}]" & vbCrLf & "Btl[{4}]" & vbCrLf & "Val[{5}]", .Y, .M, .D, B.NRec, B.TotBtl, B.TotVal)
End With
End Function

'== X
'-- Set Cpy1/2
Private Sub XVisCpy1()
SetFrmCtlnnVis XFrm, "xCpyToPth1 Cmd_Opn_CpyToPth1 Cmd_Sel_CpyToPth1", IsCpy1Pm
End Sub

Private Sub XVisCpy2()
SetFrmCtlnnVis XFrm, "xCpyToPth2 Cmd_Opn_CpyToPth2 Cmd_Sel_CpyToPth2", IsCpy2Pm
End Sub

Private Function XFrm() As Form_Rpt: Set XFrm = Form_Rpt: End Function

'------------------------------------
Private Sub XChkCpy12():
ChkTrue IsCpy1Pm Or IsCpy2Pm, "Please enable copy to Path"
If IsCpy1Pm Then ChkPthExist CpyToPth1Pm, CSub
If IsCpy2Pm Then ChkPthExist CpyToPth2Pm, CSub
End Sub
