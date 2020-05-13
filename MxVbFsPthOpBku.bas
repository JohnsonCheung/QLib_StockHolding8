Attribute VB_Name = "MxVbFsPthOpBku"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsPthOpBku."
#If Doc Then
'Bku:Cml   #Backup# used as verb
'Bu:Cml    #Backup# used as adjustive
'Fsi:Cml   #FileSystem-Item# Ffn or Pth
'P:Cml     #CurPj#
'C:Cml     #Cur#
'Ffn:Cml   #Full-File-Name#
'Pth:Cml   #Path#   A string of full path, optionally having path-separator as last char, which is preferred.
#End If
Function BkuFfn$(Ffn, Optional Msg$ = "Bku")
Const CSub$ = CMod & "BkuFfn"
Dim Tmpn$:       Tmpn = TmpNm
Dim TarFfn$:   TarFfn = BkFfn(Ffn)
Dim MsgFfn$:   MsgFfn = Pth(TarFfn) & "Msg.txt"
Dim MsgiFfn$: MsgiFfn = ParPthzFfn(TarFfn) & "MsgIdx.txt"
Dim Msgi$:       Msgi = "#" & Tmpn & vbTab & Msg & vbCrLf
:                       CpyFfn Ffn, TarFfn      ' <==
:                       WrtStr Msgi, MsgFfn     ' <==
:                       AppStr Msgi, MsgiFfn    ' <==
:                       BkuFfn = TarFfn
:                       Inf CSub, "File is Backuped", "As-file", TarFfn
End Function

Function BkPth$(Ffn) ' :Pth #Backup-Path# ! The path used backuping @Ffn
BkPth = AddFdrEns(BkHom(Ffn), TmpNm)
End Function

Function BkFfn$(Ffn)
BkFfn = BkPth(Ffn) & Fn(Ffn)
End Function

Function BkHom$(Ffn)
ChkFfnExist Ffn, "BkHom"
BkHom = EnsPth(AssPth(Ffn) & ".Backup")
End Function

Function LasBkFfn$(Ffn)
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrAyzIsInst(H)
Dim Fdr$: Fdr = MaxEle(F)
LasBkFfn = H & Fdr & "\" & Fn(Ffn)
End Function

Function BkFfnAy(Ffn) As String()
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrAyzIsInst(H)
Dim Fn1$: Fn1 = Fn(Ffn)
Dim Fdr: For Each Fdr In Itr(F)
    Dim IFfn$: IFfn = H & Fdr & "\" & Fn1
    If HasFfn(IFfn) Then
        PushI BkFfnAy, IFfn
    End If
Next
End Function

Function BkRoot$(Pth)
BkRoot = AddFdr(Pth, ".Bku")
End Function
