VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_Rpt."
Private Function xD() As Byte: xD = Me.DD: End Function
Private Function xM() As Byte: xM = Me.MM: End Function
Private Function xY() As Byte: xY = Me.YY: End Function
Private Function xYmd() As Ymd: xYmd = Ymd(xY, xM, xD): End Function
Private Function xYYMMDD&(): xYYMMDD = YYMMDD(xY, xM, xD): End Function
Private Sub Cmd_Exit_Click():             DoCmd.Close:                    End Sub

Private Sub Cmd_Opn_MacauRatePrm_Click()
DoCmd.OpenForm "MacauCalcRate"
End Sub

Private Sub Cmd_Opn_ReadMe_Click()
OpnFrm "Tx_Rpt_ReadMe"
End Sub

Private Sub Cmd_Opn_SHldOPth_Click(): BrwPth MB52OPth:  End Sub
Private Sub Cmd_Opn_MB52OPth_Click(): BrwPth MB52OPth: End Sub

Private Sub Cmd_Rfh_Report_Click(): TbReport_Rfh: Requery: End Sub

Private Sub Cmd_Sel_IPth_Click():         SelMB52InpPth: End Sub
Private Sub Cmd_Opn_IPth_Click():         BrwPth MB52IPthPm:         End Sub

Private Sub Cmd_Ena_CpyToPth1_Click():    TglIsCpy1Pm: End Sub
Private Sub Cmd_Ena_CpyToPth2_Click():    TglIsCpy2Pm: End Sub
Private Sub Cmd_Opn_CpyToPth1_Click():    BrwPth CpyToPth1Pm:      End Sub
Private Sub Cmd_Opn_CpyToPth2_Click():    BrwPth CpyToPth2Pm:      End Sub
Private Sub Cmd_Sel_CpyToPth1_Click():    Sel_CpyToPth1: End Sub
Private Sub Cmd_Sel_CpyToPth2_Click():    Sel_CpyToPth2: End Sub

Private Sub Cmd_Rec_LoadMB52_Click():     LoadMB52 xYmd: Requery: End Sub
Private Sub Cmd_Rec_Opn_GitIFx_Click():   OpnFx GitIFx(xYmd): End Sub
Private Sub Cmd_Rec_Opn_GitITbl_Click():  LnkAndOpnGitIFx xYmd: End Sub
Private Sub Cmd_Rec_OpnFx_Click():        OpnFx MB52IFx(xYmd):     End Sub
Private Sub Cmd_Rec_OpnGitIFx_Click():    OpnFx GitIFx(xYmd): End Sub
Private Sub Cmd_Rec_OpnGitITbl_Click():   LnkAndOpnGitIFx xYmd: End Sub
Private Sub Cmd_Rec_OpnMB52IFx_Click():   OpnFx MB52IFx(xYmd): End Sub
Private Sub Cmd_Rec_OpnMB52ITbl_Click():  LnkAndOpnMB52IFx xYmd: End Sub
Private Sub Cmd_Rec_RptMB52_Click():      RptMB52 xYmd: End Sub
Private Sub Cmd_Rec_RptStkHld_Click():    RptSHld xYmd:           End Sub
Private Sub Cmd_Rec_Clr_Click(): DltMB52Rec xYmd: Requery: End Sub
Private Sub Cmd_Rec_LoadGit_Click():      LoadGit xYmd: Requery: End Sub

Private Sub Cmd_Rec_CpyMB52To_Click(): CpyMB52To xYmd: End Sub
Private Sub Cmd_Rec_CpySHldTo_Click(): CpySHldTo xYmd: End Sub

Private Sub CmdMsg_Click()
ShwSts
End Sub

Private Sub Form_Load()
VisCpy12
ClrSts
End Sub
