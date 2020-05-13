Attribute VB_Name = "MxIdeCtlNoPm"
Option Explicit
Option Compare Text

Function IdeBarNy() As String(): IdeBarNy = Itn(IdeBars): End Function
Function NWin&(): NWin = CVbe.Windows.Count: End Function
Function NVisWin&(): NVisWin = NItpTrue(CVbe.Windows, "Visible"): End Function
Function IdeBars() As Office.CommandBars: Set IdeBars = BarszV(CVbe): End Function
Function WinCapy() As String(): WinCapy = SyzItp(CVbe.Windows, "Caption"): End Function
Function VisWinCapy() As String(): VisWinCapy = SyzOyP(VisWiny, "Caption"): End Function

Function VisWiny() As VBIde.Window()
Dim W As VBIde.Window: For Each W In CVbe.Windows
    If W.Visible Then PushObj VisWiny, W
Next
End Function

Function SampBtnSpec() As String()
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
SampBtnSpec = XX
End Function


