VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LoadPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_LoadPH."
Option Base 0
Public xLvl As Byte
Private Sub SetFilter(Lvl As Byte)
Me.RecordSource = "Select * from ProdHierarchy" & PH_Where(Lvl, Nz(Me.Cmd_Filter_Cur.Value, False), Nz(Me.Cmd_Filter_Hst, False)) & " Order by Srt,Des"
xLvl = Lvl
Cmd_Filter_L1.Value = False
Cmd_Filter_L2.Value = False
Cmd_Filter_L3.Value = False
Cmd_Filter_L4.Value = False
Select Case xLvl
Case 1: Cmd_Filter_L1.Value = True
Case 2: Cmd_Filter_L2.Value = True
Case 3: Cmd_Filter_L3.Value = True
Case 4: Cmd_Filter_L4.Value = True
End Select
Me.Requery
End Sub

Private Sub Cmd_Filter_Cur_Click()
If Me.Cmd_Filter_Cur.Value And Me.Cmd_Filter_Hst.Value Then
    Me.Cmd_Filter_Hst.Value = False
End If
SetFilter xLvl
End Sub

Private Sub Cmd_Filter_Hst_Click()
If Me.Cmd_Filter_Hst.Value And Me.Cmd_Filter_Cur.Value Then
    Me.Cmd_Filter_Cur.Value = False
End If
SetFilter xLvl
End Sub

Private Sub Cmd_Filter_L1_Click()
If Me.Cmd_Filter_L1.Value Then SetFilter 1 Else Me.Cmd_Filter_L1.Value = -1
End Sub

Private Sub Cmd_Filter_L2_Click()
If Me.Cmd_Filter_L2.Value Then SetFilter 2 Else Me.Cmd_Filter_L2.Value = -1
End Sub

Private Sub Cmd_Filter_L3_Click()
If Me.Cmd_Filter_L3.Value Then SetFilter 3 Else Me.Cmd_Filter_L3.Value = -1
End Sub

Private Sub Cmd_Filter_L4_Click()
If Me.Cmd_Filter_L4.Value Then SetFilter 4 Else Me.Cmd_Filter_L4.Value = -1
End Sub

Private Sub Cmd_Sel_InpPth_Click()
SelCFxPm "PH_InpFx", "Select Product Hierarchy Xlsx file", Me.xFx
End Sub

Private Sub Cmd_Srt_Click()
SavRec
RfhTbPH_FldSno
RfhTbPH_FldSrt
Me.OrderBy = "Srt,Des"
Me.OrderByOn = True
End Sub

Private Sub Cmd_Exit_Click()
RfhTbPH_FldSno
RfhTbPH_FldSrt
DoCmd.Close
End Sub
Private Sub Cmd_Imp_Click():        LoadPH Me:                       End Sub
Private Sub Cmd_Opn_InpPth_Click(): BrwPth Pth(PHIFx): End Sub
Private Sub Cmd_Opn_InpFx_Click():  OpnFx PHIFx:              End Sub
Private Sub Cmd_Sel_InpFx_Click():  SelCFxPm "PHPrm_InpFx", , Me.xFx:         End Sub

Private Sub Form_Open(Cancel As Integer)
RfhTbPH_Fld_WithOHxxx
Me.Cmd_Filter_Cur.Value = False
Me.Cmd_Filter_Hst.Value = True
SetFilter 4
Me.Filter = ""
DoCmd.Maximize
End Sub

Private Sub Sno_BeforeUpdate(Cancel As Integer)
If Me.Lvl < xLvl Then Cancel = True: MsgBox "Cannot change": Me.Undo: Exit Sub
Me.DteUpd.Value = Now
End Sub


Function IsParentLvl(Lvl) As Boolean
IsParentLvl = xLvl > Lvl
End Function
