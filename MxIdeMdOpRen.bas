Attribute VB_Name = "MxIdeMdOpRen"
Option Explicit
Option Compare Text
Const CNs$ = "Md.Op"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdRen."

Sub Ren(FmCmpn, ToNm)
Const CSub$ = CMod & "RenTo"
If HasCmpzP(CPj, ToNm) Then Inf CSub, "CmpToNm exist", "ToNm", ToNm: Exit Sub
Cmp(FmCmpn).Name = ToNm
End Sub

Sub RenM(NewCmpn)
CCmp.Name = NewCmpn
End Sub

Sub RenMdPfx(FmPfx$, ToPfx$, Optional Pj As VBProject)
Dim P As VBProject: Set P = DftPj(Pj)
Dim C As VBComponent
For Each C In P.VBComponents
    If HasPfx(C.Name, FmPfx) Then
        RenMd C.CodeModule, RplPfx(C.Name, FmPfx, ToPfx)
    End If
Next
End Sub

Sub RenMd(M As CodeModule, NewNm$)
If HasMd(PjzM(M), NewNm) Then
    Debug.Print "New mdn[" & NewNm & "] exist, cannot rename"
    Exit Sub
End If
M.Parent.Name = NewNm
End Sub
Sub RenMdByRmvPfx(Pj As VBProject, Pfx$)
Dim C As VBComponent
For Each C In Pj.VBComponents
    If HasPfx(C.Name, Pfx) Then
        RenCmp C, RmvPfx(C.Name, Pfx)
    End If
Next
End Sub
