Attribute VB_Name = "MxIdeCur"
Option Explicit
Option Compare Text
Const CNs$ = "Cur.Ide"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeCur."

Function CLnozM&(M As CodeModule): CLnozM = RRCCzPne(M.CodePane).R1: End Function
Function CCmp() As VBComponent: Set CCmp = CMd.Parent: End Function
Function CMd() As CodeModule: Set CMd = CPne.CodeModule: End Function
Function CLnoM&(): CLnoM = CLnozM(CMd): End Function
Function CMdn(): CMdn = CCmp.Name: End Function
Function CMdDn$(): CMdDn = MdDn(CMd): End Function
Function CWin() As VBIde.Window: Set CWin = CPne.Window: End Function
Function CPne() As VBIde.CodePane: Set CPne = CVbe.ActiveCodePane: End Function
Function CMthlno&(): CMthlno = MthlnozCLno(CMd, CLnoM): End Function
Function CVbe() As Vbe: Set CVbe = Application.Vbe: End Function
Function PthP$(): PthP = PthzP(CPj): End Function
Function PjfnP$(): PjfnP = Fn(CPj.FileName): End Function
Function CPj() As VBProject: Set CPj = CVbe.ActiveVBProject: End Function
Function CMainPj() As VBProject: Set CMainPj = MainPj(Acs): End Function
Function CPjn$(): CPjn = CPj.Name: End Function
Function CPjf$(): CPjf = PjfP: End Function ' Current project file
Function PjfP$(): PjfP = Pjf(CPj): End Function ' Project file of current project
Function MthDnM$(): MthDnM = MthDnzM(CMd, CMthln): End Function
Function PthzP$(P As VBProject): PthzP = Pth(P.FileName): End Function
Function MainPj(A As Access.Application) As VBProject ' main project
Dim P As VBProject: For Each P In A.Vbe.VBProjects
    If A.CurrentDb.Name = P.FileName Then Set MainPj = P: Exit Function
Next
Imposs "MainPj", FmtQQ("Each @Acs with CurrentDb should have a MainPj.  @Acs[?]", A.CurrentDb.Name)
End Function



