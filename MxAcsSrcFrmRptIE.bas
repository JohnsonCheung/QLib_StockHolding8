Attribute VB_Name = "MxAcsSrcFrmRptIE"
Option Compare Text
Option Explicit
Const CNs$ = "Acs"
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsSrcFrmRptIE."
':Nwn: :Cml #New-Name#
Sub ExpAllFrm(A As Access.Application, Pth$)
Dim F: For Each F In Itr(FrmNy(A))
    ExpFrm A, F, Pth
Next
End Sub

Sub ImpAllFrm(A As Access.Application, Pth$)
Dim F: For Each F In Itr(FfnAy(Pth, "*.frm"))
    ImpFrm A, F
Next
End Sub

Sub ExpAllCRpt(Pth$): ExpAllRpt Acs, Pth: End Sub
Sub ExpAllCFrm(Pth$): ExpAllFrm Acs, Pth: End Sub
Sub ImpCFrm(FrmFfn, Optional Nwn$): ImpFrm Acs, FrmFfn, Nwn$: End Sub
Sub ImpCRpt(RptFfn, Optional Nwn$): ImpRpt Acs, RptFfn, Nwn$: End Sub
Sub ExpCFrm(F, Pth$): ExpFrm Acs, F, Pth: End Sub
Sub ExpCRpt(R, Pth$): ExpFrm Acs, R, Pth: End Sub

Sub ExpAllRpt(A As Access.Application, Pth$)
Dim F: For Each F In Itr(RptNy(A))
    ExpRpt A, F, Pth
Next
End Sub

Sub ImpAllRpt(A As Access.Application, Pth$)
Dim F: For Each F In Itr(FfnAy(Pth, "*.rpt"))
    ImpRpt A, F
Next
End Sub


Sub ImpRpt(A As Access.Application, RptFfn, Optional Nwn$): A.LoadFromText acReport, Nm_(Nwn, RptFfn), RptFfn: End Sub
Sub ImpFrm(A As Access.Application, FrmFfn, Optional Nwn$): A.LoadFromText acForm, Nm_(Nwn, FrmFfn), FrmFfn:   End Sub

Sub ExpFrm(A As Access.Application, F, Pth$): A.SaveAsText acForm, F, EnsPthSfx(Pth) & F & ".frm": End Sub
Sub ExpRpt(A As Access.Application, R, Pth$): A.SaveAsText acReport, R, EnsPthSfx(Pth) & R & ".rpt": End Sub

Private Function Nm_$(Nwn$, Ffn)
If Nwn <> "" Then Nm_ = Nwn: Exit Function
Nm_ = Fnn(Ffn)
End Function
