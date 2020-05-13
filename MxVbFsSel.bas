Attribute VB_Name = "MxVbFsSel"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "InterAct"
Const CMod$ = CLib & "MxVbFsSel."


Sub SetTxtbSelPth(A As Access.TextBox)
Dim R$
R = SelPth(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub


Function SelPth$(Optional Pth$, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = Pth
    .Show
    If .SelectedItems.Count = 1 Then
        SelPth = EnsPthSfx(.SelectedItems(1))
    End If
End With
End Function

Private Sub SelPth__Tst()
GoTo Z
Z:
MsgBox SelFfn("C:\")
End Sub

Function SelFx$(Optional DftFx$, Optional SpecDes$ = "Select a Xlsx file")
SelFx = SelFfn(DftFx, "*.xlsx", SpecDes)
End Function

Function SelFfn$(Optional Ffn$, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With Application.FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    If HasFfn(Ffn) Then .InitialFileName = Ffn
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        SelFfn = .SelectedItems(1)
    End If
End With
End Function
