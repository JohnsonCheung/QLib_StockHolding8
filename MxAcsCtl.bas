Attribute VB_Name = "MxAcsCtl"
Option Explicit
Option Compare Text
Const CNs$ = "Acs.Ctl"
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsCtl."

Function CvAcsCtl(A) As Access.Control
Set CvAcsCtl = A
End Function

Function CvAcsBtn(A) As Access.CommandButton
Set CvAcsBtn = A
End Function

Function CvAcsTgl(A) As Access.ToggleButton
Set CvAcsTgl = A
End Function

Sub SetAcsPrp(C As Access.Control, P, V)
Dim I As AccessObjectProperty: For Each P In C.Properties
    If I.Name = P Then
        I.Value = V
        Exit Sub
    End If
Next
End Sub
Sub SetTabStop(A As Access.Form, OnOff As Boolean)
Dim C As Access.Control: For Each C In A.Controls
    SetAcsPrp C, "TabStop", OnOff
Next
End Sub
Function CFrm() As Access.Form
If CurrentObjectType = acForm Then
    Set CFrm = Access.Forms(CurrentObjectName)
End If
End Function
Function HasCtl(F As Access.Form, Nm$) As Boolean
Dim C As Access.Control: For Each C In F.Controls
    If C.Name = Nm Then HasCtl = True: Exit Function
Next
End Function

Function IsFrmOpn(FrmNm$): IsFrmOpn = CurrentProject.AllForms(FrmNm).CurrentView = acCurViewFormBrowse: End Function

Sub SetFrmCtlnnVis(F As Access.Form, Ctlnn$, Vis As Boolean)
Dim N: For Each N In Split(Ctlnn)
    F.Controls(N).Visible = Vis
Next
End Sub
