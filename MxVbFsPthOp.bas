Attribute VB_Name = "MxVbFsPthOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsPthOp."

Sub VcPth(Pth)
If NoPth(Pth) Then Exit Sub
Shell FmtQQ("Code.cmd ""?""", Pth), vbMaximizedFocus
End Sub

Sub BrwPth(Pth)
If NoPth(Pth) Then Exit Sub
ShellMax FmtQQ("Explorer ""?""", Pth)
End Sub

Sub DltPth(Pth)
ChkPthExist Pth, "DltPth"
RmDir Pth
End Sub

Private Sub DltAllPthFil__Tst()
DltAllPthFil TmpRoot
End Sub

Sub DltAllPthFil(Pth)
If NoPth(Pth) Then Exit Sub
Dim F
For Each F In Itr(Ffny(Pth))
   DltFfn F
Next
End Sub

Sub DltEmpPthR(Pth)
Dim Ay$(), I, J%
Lp:
    J = J + 1: If J > 10000 Then Stop
    Dim Dlt As Boolean: Dlt = False
    For Each I In Itr(EmpPthAyR(Pth))
        DltPthSilent I
        Dlt = True
    Next
    If Dlt Then GoTo Lp
End Sub
Sub DltPthSilent(Pth)
On Error Resume Next
RmDir Pth
End Sub

Sub DltAllEmpFdr(Pth)
Dim S: For Each S In Itr(SubPthAy(Pth))
   DltPthIfEmp S
Next
End Sub

Sub DltPthIfEmp(Pth)
If IsEmpPth(Pth) Then DltPth Pth
End Sub

Sub RenPthAddFdrPfx(Pth, Pfx)
RenPth Pth, AddFdrPfx(Pth, Pfx)
End Sub

Sub RenPth(Pth, NewPth)
Fso.GetFolder(Pth).Name = NewPth
End Sub

Private Sub DltEmpSubDir__Tst()
DltAllEmpFdr TmpHom
End Sub
