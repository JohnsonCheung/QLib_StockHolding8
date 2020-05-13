VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LoadFc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_LoadFc."
Private Property Get xY() As Byte:  xY = Nz(Me.VerYY.Value, 0):  End Property
Private Property Get xM() As Byte:  xM = Nz(Me.VerMM.Value, 0):  End Property
Private Property Get xWFx$():      xWFx = FcWFx(xStmYM):  End Property
Private Property Get xIFx$():      xIFx = FcIFx(xStmYM):  End Property
Private Property Get XStm$():       XStm = Nz(Me.Stm.Value, ""): End Property
Private Property Get xStmYM() As StmYM: xStmYM = StmYM(XStm, xY, xM): End Property

Private Sub Cmd_Opn_WFx_Click(): OpnFx xWFx:  End Sub
Private Sub Cmd_Opn_IFx_Click(): OpnFx xIFx: End Sub

Private Sub Cmd_Opn_IPth_Click(): OpnFcIPth: End Sub
Private Sub Cmd_Opn_WPth_Click(): OpnFcWPth: End Sub

Private Sub Cmd_Clr_Click():  ClrFc xStmYM:  End Sub
Private Sub Cmd_Load_Click(): LoadFc xStmYM: End Sub
Private Sub Cmd_Opn_OupPth_Click(): BrwPth FcOPth: End Sub

Private Sub Cmd_Rfh_Click()
RfhTbFc_FmFcIPth
End Sub

Private Sub Cmd_Rpt_Fc_Click()
RptFc YM(xY, xM)
End Sub

Private Sub Form_Open(Cancel As Integer):
RfhTbFc_FmFcIPth
DoCmd.Maximize
End Sub
Private Sub Cmd_Exit_Click():             DoCmd.Close:    End Sub

Sub OpnFcIPth()
BrwPth FcIPth
End Sub

Sub OpnFcWPth()
BrwPth FcWPth
End Sub
