Attribute VB_Name = "MxDtaS32"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDtaS32."
Const CNs$ = "Fs"

Function S32HdrzFdr(F As Shell32.Folder) As String()
':S32: :Shell32 ! #Shell32#
Dim J%: Do
    Dim M$: M = F.GetDetailsOf(Null, J)
    If M = "" Then Exit Function
    PushI S32HdrzFdr, M
    LoopTooMuch "S32HdrzFdr", J
Loop
End Function

Function S32Fdr(Pth) As Shell32.Folder
Dim S As New Shell
Dim F As Shell32.Folder
Set S32Fdr = S.Namespace(RmvPthSfx(Pth))
End Function

Private Sub S32Hdr__Tst()
Dmp S32Hdr("C:\Users\User")
End Sub

Function S32Hdr(Pth) As String()
S32Hdr = S32HdrzFdr(S32Fdr(Pth))
End Function
Function S32ItmsDr(F As Shell32.Folder, I As Shell32.FolderItem2, HdrUB&) As Variant()
Dim J%: For J = 0 To HdrUB
    PushI S32ItmsDr, F.GetDetailsOf(I, J)
Next
End Function

Private Sub S32ItmsDrs__Tst()
BrwDrsN S32ItmsDrs("C:\Users\user\Documents\Projects\Vba\QLib\QLib.xlam.res")
End Sub

Function S32ItmsDrs(Pth) As Drs
Dim F As Shell32.Folder: Set F = S32Fdr(Pth)
Dim Hdr$(): Hdr = S32HdrzFdr(F)
Dim HdrUB&: HdrUB = UB(Hdr)
Dim ODy() ': Dim F As Shell32.Folder, HdrUB&
    Dim A_IItm As Shell32.FolderItem2: For Each A_IItm In F.Items
        Dim A_Dr(): A_Dr = S32ItmsDr(F, A_IItm, HdrUB)
        PushI ODy, A_Dr
    Next
S32ItmsDrs = Drs(Hdr, ODy)
End Function
