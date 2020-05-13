Attribute VB_Name = "MxIdeCtlStd"
Option Explicit
Option Compare Text
Const CNs$ = "Ide.Ctl"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeCtlBar."
Function IdeBtnEdtClr() As Office.CommandBarButton: Set IdeBtnEdtClr = FstCaption(IdePopEdt.Controls, "C&lear"): End Function
Function IdeBarStd() As Office.CommandBar: Set IdeBarStd = IdeBars("Standard"): End Function
Function IdeBarMnu() As CommandBar: Set IdeBarMnu = IdeBarMnuzV(CVbe): End Function
Function IdeWinImm() As VBIde.Window: Set IdeWinImm = FstWin(vbext_wt_Immediate): End Function
Function IdeWinLcl() As VBIde.Window: Set IdeWinLcl = FstWin(vbext_wt_Locals): End Function
Function IdeWinBrwObj() As VBIde.Window: Set IdeWinBrwObj = FstWin(vbext_wt_Browser): End Function
Function IdeBtnSelAll() As Office.CommandBarButton: Set IdeBtnSelAll = FstCaption(IdePopEdt.Controls, "Select &All"): End Function


Private Sub IdeBarMnu__Tst()
Dim A As CommandBar
Set A = IdeBarMnu
Stop
End Sub

Function IdeBtnNxtStmt() As Office.CommandBarButton: Set IdeBtnNxtStmt = IdePopDbg.Controls("Show Next Statement"): End Function
Function IdeBtnTileH() As Office.CommandBarButton: Set IdeBtnTileH = IdePopWin.Controls("Tile &Horizontally"): End Function
Function IdeBtnTileV() As Office.CommandBarButton: Set IdeBtnTileV = IdePopWin.Controls("Tile &Vertically"): End Function
Function IdeBtnSav() As Office.CommandBarButton: Set IdeBtnSav = IdeBtnSavzV(CVbe): End Function
Function IdeBtnXls() As Office.CommandBarControl: Set IdeBtnXls = IdeBarStd.Controls(1): End Function
Function IdeBtnCompile() As Office.CommandBarButton: Set IdeBtnCompile = IdePopDbg.CommandBar.Controls(1): End Function
Function IdePopWin() As CommandBarPopup: Set IdePopWin = IdeBarMnu.Controls("Window"): End Function
Function IdePopDbg() As CommandBarPopup: Set IdePopDbg = IdeBarMnu.Controls("Debug"): End Function
Function IdePopEdt() As CommandBarPopup: Set IdePopEdt = FstCaption(IdeBarMnu.Controls, "&Edit"): End Function

Function IdeBtnSavzV(A As Vbe) As CommandBarButton
Dim I As CommandBarControl: For Each I In IdeBarStdzV(A).Controls
'    Debug.Print I.Caption
    If HasPfx(I.Caption, "&Save") Then Set IdeBtnSavzV = I: Exit Function
Next
End Function

