Attribute VB_Name = "MxIdeCltWinOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeCltWinOp."

Sub ClsAllWin()
Dim W As VBIde.Window: For Each W In CVbe.Windows
    If Not IsEqObj(W, IdeWinImm) Then ClsWin W
Next
TileV
End Sub

Sub VisWin(W As VBIde.Window)
W.Visible = True
End Sub

Sub ClsWinExlAp(ParamArray ExlWinAp())
Dim Av(): Av = ExlWinAp
Dim I: For Each I In Itr(VisWiny)
    Dim W As VBIde.Window: Set W = I
    If Not HasObj(Av, W) Then
        ClsWin W
    Else
        VisWin W
    End If
Next
End Sub

Sub ShwDbg()
ClsWinExlAp IdeWinImm, IdeWinLcl, CWin
DoEvents
TileV
End Sub

Sub ClrImm()
Dim W As VBIde.Window
DoEvents
With IdeWinImm
    .SetFocus
    .Visible = True
End With
DoEvents
SndKeys "^{HOME}^+{END}"
'SndKeys "{DEL}" '<-- it does not work?
'DoEvents
End Sub

Sub ClsWin(W As VBIde.Window)
If W.Visible Then W.Close
End Sub

Sub ClsWinE(Exl As VBIde.Window)
'Do : Cls win ept cur md @@
Dim W As VBIde.Window: For Each W In CVbe.Windows
    If Not IsEqObj(Exl, W) Then ClsWin W
Next
TileV
End Sub

Sub ClrWin(A As VBIde.Window)
DoEvents
IdeBtnSelAll.Execute
DoEvents
SendKeys " "
IdeBtnEdtClr.Execute
End Sub
