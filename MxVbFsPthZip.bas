Attribute VB_Name = "MxVbFsPthZip"
Option Compare Text
Option Explicit
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxVbFsPthZip."

Function ZipPth$(Z7Db As Database, FmPth, ToPth)
If IsPmEr(FmPth, ToPth) Then Exit Function
Dim P$:             P = AddFdrEns(Pth(Z7Db.Name), "ZipPthWrking")
Dim Fexe$:       Fexe = P & "z7.exe"
Dim Fcmd$:       Fcmd = P & "Zip.Cmd"
Dim Foup$:       Foup = Fdr(FmPth) & "(" & Format(Now, "YYYY-MM-DD HH-MM") & " " & CUsr & ").zip"
Dim Fcxt$:       Fcxt = ZipFcxt(P, FmPth, ToPth, Foup)
Dim ShellStr$: ShellStr = FmtQQ("Cmd.Exe /C ""?""", Fcmd)
                          Expz7 Z7Db, Fexe
                          WrtStr Fcxt, Fcmd
                          Shell Fcmd, vbMaximizedFocus
                 ZipPth = EnsPthSfx(ToPth) & Foup
End Function
Function CUsr$()
CUsr = Environ("USER")
End Function
Private Sub Exp7z__Tst()
':Fexe: :Ffn ! #Exe-Ffn#
Dim Z7Db As Database, Fexe$: Set Z7Db = CurrentDb: Fexe = "C:\users\user\documents\projects\vba\Bkupth\7z.exe"
End Sub

Sub Expz7(Z7Db As Database, Fexe$)
If HasFfn(Fexe) Then Exit Sub
Dim R As DAO.Recordset: Set R = Z7Db.TableDefs("7z").OpenRecordset
Dim R2 As DAO.Recordset2: Set R2 = R.Fields("7z").Value
Dim F2 As DAO.Field2: Set F2 = R2.Fields("FileData")
Dim NoExt$: NoExt = RmvExt(Fexe)
DltFfnIf NoExt
F2.SaveToFile NoExt
Name NoExt As Fexe
End Sub

Private Function ZipFcxt$(WrkPth, FmPth, ToPth, Foup$)
':Fcxt: :Lines ! #File-Context#
Dim O$()
Push O, FmtQQ("Cd ""?""", WrkPth)
Push O, FmtQQ("z7 a ""?"" ""?""", Foup, FmPth)
Push O, FmtQQ("Copy ""?"" ""?""", Foup, ToPth)
Push O, FmtQQ("Del ""?""", Foup)
Push O, "Pause"
ZipFcxt = JnCrLf(O)
End Function

Private Function IsPmEr(FmPth, ToPth) As Boolean
IsPmEr = True
If NoPth(ToPth) Then SetMainMsg "To Path not found: " & ToPth: Exit Function
If NoPth(FmPth) Then SetMainMsg "From Path not found: " & FmPth: Exit Function
IsPmEr = False
End Function

Private Sub ZipPth__Tst()
Dim A$: A = TmpPthi
ZipPth CurrentDb, PthP, A
End Sub
