Attribute VB_Name = "MxIdeCtlOp"
Option Compare Text
Option Explicit
Const CNs$ = "Ctl.Op"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeCtlOp."
Sub NxtStmt()
Const CSub$ = CMod & "JmpNxtStmt"
With IdeBtnNxtStmt
    If Not .Enabled Then
        'Msg CSub, "BoJmpNxtStmt is disabled"
        Exit Sub
    End If
    .Execute
End With
End Sub

Sub TileH(): IdeBtnTileH.Execute: End Sub
Sub MaxiImm(): IdeWinImm.WindowState = vbext_ws_Maximize: End Sub
Sub TileV(): IdeBtnTileV.Execute: End Sub

Sub CompileP(): CompilezP CPj: End Sub
Sub Compile(Pjn$)
JmpPj Pj(Pjn)
IdeBtnCompile.Execute
End Sub


Sub CompilezP(P As VBProject)
JmpPj P
With IdeBtnCompile
    If .Enabled Then
        .Execute
        Debug.Print P.Name, "<--- Compiled"
    Else
        Debug.Print P.Name, "already Compiled"
    End If
End With
IdeBtnTileV.Execute
IdeBtnSav.Execute
End Sub

Sub CompilezV(A As Vbe): ItrDo A.VBProjects, "CompilezP": End Sub

Sub ChkCompileBtnGood(Pjn$, Fun$)
Dim Act$, Ept$
Act = IdeBtnCompile.Caption
Ept = "Compi&le " & Pjn
If Act <> Ept Then Thw CSub, "Cur CompileBtn.Caption <> Compi&le {Pjn}", "Compile-Btn-Caption Pjn Ept-Btn-Caption", Act, Pjn, Ept
End Sub

Private Sub CompilezP__Tst()
CompilezP CPj
End Sub

Sub DltClr(A As CommandBar)
Dim I: For Each I In Itr(OyzItr(A.Controls))
    CvCtl(I).Delete
Next
End Sub

Sub DltBar(BarNm$): IdeBars(BarNm).Delete: End Sub

Sub EnsBtns(BarBtnccAy$())
Dim I: For Each I In Itr(BarBtnccAy)
    EnsBarBtncc I
Next
End Sub

Sub EnsBarBtncc(BarBtncc)
Dim L$: L = BarBtncc
EnsBtnzCC EnsBar(ShfTerm(L)), L
End Sub

Sub RmvBarNy(BarNy$())
Dim IBar: For Each IBar In BarNy
    If HasBar(IBar) Then
        If Not IdeBar(IBar).BuiltIn Then
            IdeBar(IBar).Delete
        End If
    End If
Next
End Sub

Function EnsBar(BarNm$) As CommandBar
If HasBar(BarNm) Then
    Set EnsBar = IdeBars(BarNm)
Else
    Set EnsBar = IdeBars.Add(BarNm)
End If
EnsBar.Visible = True
End Function

Sub EnsBtnzCC(Bar As CommandBar, BtnCapcc$)
Dim BtnCap
For Each BtnCap In Termy(BtnCapcc)
    EnsBtnzC Bar, BtnCap
Next
End Sub

Function HasBtn(Bar As CommandBar, BtnCap) As Boolean
Dim C As CommandBarControl
For Each C In Bar.Controls
    If C.Type = msoControlButton Then
        If C.Caption = BtnCap Then HasBtn = True: Exit Function
    End If
Next
End Function

Sub EnsBtnzC(Bar As CommandBar, BtnCap)
If HasBtn(Bar, BtnCap) Then Exit Sub
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = BtnCap
B.Style = msoButtonCaption
End Sub

Sub AddBtn(Bar As CommandBar, BtnCap)
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = BtnCap
B.Style = msoButtonCaption
End Sub

Private Function SampToolBarSpec() As String()
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
SampToolBarSpec = XX  '*Spec
Erase XX
End Function

Function BtnSpec(ToolBarSpec$()) As String()
BtnSpec = IndLy(ToolBarSpec, "Bars")
End Function

Sub InstallIdeTools(ToolBarSpec$())
EnsBtns BtnSpec(ToolBarSpec)
'EnsMdl Md("IdeTool"), ToolClsCd
End Sub

Function ToolBarNy(ToolBarSpec$()) As String(): ToolBarNy = AmT1(BtnSpec(ToolBarSpec)): End Function
Sub RmvIdeTools(ToolBarSpec$()): RmvBarNy ToolBarNy(ToolBarSpec): End Sub
