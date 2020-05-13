Attribute VB_Name = "MxIdeCtlFun"
Option Compare Text
Option Explicit
Const CNs$ = "Ide.Ctl"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeCtlBtn."
Function BarszV(A As Vbe) As Office.CommandBars: Set BarszV = A.CommandBars: End Function
Function Capy(A As Controls) As String(): Capy = SyzItrP(A, "Caption"): End Function
Function CvCtl(A) As CommandBarControl: Set CvCtl = A: End Function
Function WinzMdn(Mdn) As VBIde.Window: Set WinzMdn = Md(Mdn).CodePane.Window: End Function
Function WinzM(M As CodeModule) As VBIde.Window: Set WinzM = M.CodePane.Window: End Function
Function FstCaption(Itr, Caption): FstCaption = FstObjByEq(Itr, "Caption", Caption): End Function
Function HasBar(BarNm) As Boolean: HasBar = HasBarzV(CVbe, BarNm): End Function
Function IsBtn(A) As Boolean: IsBtn = TypeName(A) = "CommandButton": End Function
Function CvBtn(A) As Office.CommandBarButton: Set CvBtn = A: End Function
Function CvWin(A) As VBIde.Window: Set CvWin = A: End Function
Private Sub IdePopDbg__Tst()
Dim A
Set A = IdePopDbg
Stop
End Sub

Function PnezCmpn(Cmpn$) As CodePane: Set PnezCmpn = Md(Cmpn).CodePane: End Function
Function IdeBar(BarNm) As Office.CommandBar: Set IdeBar = IdeBars(BarNm): End Function
Function IdeBarStdzV(A As Vbe) As Office.CommandBar: Set IdeBarStdzV = IdeBar("Standard"): End Function
Function IdeBarMnuzV(A As Vbe) As CommandBar: Set IdeBarMnuzV = A.CommandBars("Menu Bar"): End Function
Sub ClsWinExlMdn(ExlMdn$): ClsWinExlAp IdeWinImm, WinzMdn(ExlMdn): End Sub
Function BarNyzV(A As Vbe) As String(): BarNyzV = Itn(A.CommandBars): End Function

Function RRCCzPne(P As CodePane) As RRCC
Dim R1&, R2&, C1&, C2&
P.GetSelection R1, R2, C1, C2
RRCCzPne = RRCC(R1, R2, C1, C2)
End Function

Function FstWin(A As vbext_WindowType) As VBIde.Window: Set FstWin = FstObjByEq(CVbe.Windows, "Type", A): End Function
Function WinyzTy(T As vbext_WindowType) As VBIde.Window(): WinyzTy = IwEq(CVbe.Windows, "Type", T): End Function
Function MdnzCdWin$(CdWin As VBIde.Window): MdnzCdWin = IsBet(CdWin.Caption, " - ", " (Code)"): End Function



