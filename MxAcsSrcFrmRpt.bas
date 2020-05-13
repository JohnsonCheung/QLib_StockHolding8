Attribute VB_Name = "MxAcsSrcFrmRpt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxAcsSrcFrmRpt."
Function FrmTmpFt$(A As Access.Application, Frmn$)
Dim T$: T = TmpFfn(".frm", "AcsFrm", Frmn)
A.SaveAsText acForm, Frmn, T
FrmTmpFt = T
End Function
Function RptTmpFt$(A As Access.Application, Rptn$)
Dim T$: T = TmpFfn(".Rpt", "AcsRpt", Rptn)
A.SaveAsText acReport, Rptn, T
RptTmpFt = T
End Function

Sub BrwCFrmSrc(Frmn$): BrwFrmSrc Acs, Frmn: End Sub
Sub BrwCRptSrc(Rptn$): BrwRptSrc Acs, Rptn: End Sub
Sub VcCFrmSrc(Frmn$): VcFrmSrc Acs, Frmn: End Sub
Sub VcCRptSrc(Rptn$): VcRptSrc Acs, Rptn: End Sub

Sub BrwFrmSrc(A As Access.Application, Frmn$): BrwFt FrmTmpFt(A, Frmn): End Sub
Sub BrwRptSrc(A As Access.Application, Rptn$): BrwFt RptTmpFt(A, Rptn): End Sub
Sub VcFrmSrc(A As Access.Application, Frmn$): VcFt FrmTmpFt(A, Frmn): End Sub
Sub VcRptSrc(A As Access.Application, Rptn$): VcFt RptTmpFt(A, Rptn): End Sub
