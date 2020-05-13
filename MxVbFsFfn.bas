Attribute VB_Name = "MxVbFsFfn"
Option Compare Text
Option Explicit
Public Fso As New FileSystemObject
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsFfn."
Public Const FbExt$ = ".accdb"
Public Const FbExt1$ = ".mdb"
Public Const FbaExt$ = ".accdb"
Public Const FxaExt$ = ".xlam"
Enum eFilCprMth
    eCprEachbyt
    eCprTimSi
End Enum
Public Const PthSep$ = "\"
Function EoFfnMis(Ffn, Optional Kd$ = "File") As String()
If HasFfn(Ffn) Then Exit Function
BfrClr
BfrV FmtQQ("? not found", Kd)
BfrTab "Path : " & Pth(Ffn)
BfrTab "File : " & Fn(Ffn)
EoFfnMis = BfrLy
End Function
Function CutPth$(Ffn)
Dim P%: P = InStrRev(Ffn, PthSep)
If P = 0 Then CutPth = Ffn: Exit Function
CutPth = Mid(Ffn, P + 1)
End Function
Function FnzFfn$(Ffn)
FnzFfn = CutPth(Ffn)
End Function

Function Fn$(Ffn)
Fn = CutPth(Ffn)
End Function

Function FfnUp$(Ffn)
FfnUp = ParPth(Pth(Ffn)) & Fn(Ffn)
End Function

Function Fnn$(Ffn)
Fnn = RmvExt(Fn(Ffn))
End Function

Function RmvExt$(Ffn)
Dim B$, C$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
RmvExt = Pth(Ffn) & C
End Function

Function IsExtInAp(Ffn, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap: IsExtInAp = IsInAv(Ext(Ffn), Av)
End Function

Function IsInAv(V, Av()) As Boolean
IsInAv = HasEle(Av, V)
End Function

Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
IsInAp = HasEle(Av, V)
End Function

Function ExtzFfn$(Ffn)
ExtzFfn = Ext(Ffn)
End Function

Function Ext$(Ffn)
Dim B$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
Ext = Mid(B, P)
End Function

Function UpPth$(Pth, Optional NUp% = 1)
Dim O$: O = Pth
Dim J%
For J = 1 To NUp
    O = ParPth(O)
Next
UpPth = O
End Function

Function Pth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn, P)
End Function

Function ParPthzFfn$(Ffn)
ParPthzFfn = ParPth(Pth(Ffn))
End Function

Function IsEqFfnStr(Ffn, S$) As Boolean
Dim L&: L = Len(S)
If FileLen(Ffn) <> L Then Exit Function
Dim J&, F%
F = FnoRnd128(Ffn)
For J = 1 To NBlk(SizFfn(Ffn), 128)
    Dim P&: P = (J - 1) * 128 + 1
    If FnoBlk(F, J) <> Mid(S, P, 128) Then
        Close #F
        Exit Function
    End If
Next
Close #F
IsEqFfnStr = True
End Function

Function IsEqFfn(A, B, Optional M As eFilCprMth = eFilCprMth.eCprEachbyt) As Boolean
Const CSub$ = CMod & "IsEqFfn"
ChkFfnExist A, CSub, "Fst File"
If A = B Then Thw CSub, "Fil A and B are eq name", "A", A
ChkFfnExist B, CSub, "Snd File"
If Not IsSamTimSi(A, B) Then Exit Function
If M = eCprTimSi Then
    IsEqFfn = True
    Exit Function
End If
Dim J&, F1%, F2%
F1 = FnoRnd128(A)
F2 = FnoRnd128(B)
For J = 1 To NBlk(SizFfn(A), 128)
    If FnoBlk(F1, J) <> FnoBlk(F2, J) Then
        Close #F1, F2
        Exit Function
    End If
Next
Close #F1, F2
IsEqFfn = True
End Function

Function IsSamTimSi(Ffn1, Ffn2) As Boolean
If DtezFfn(Ffn1) <> DtezFfn(Ffn2) Then Exit Function
If Not IsSamzSi(Ffn1, Ffn2) Then Exit Function
IsSamTimSi = True
End Function

Function IsSamzSi(Ffn1, Ffn2) As Boolean
IsSamzSi = SizFfn(Ffn1) = SizFfn(Ffn2)
End Function

Function MsgSamFfn(A, B, Si&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Si
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
MsgSamFfn = O
End Function

Private Sub FfnBlk__Tst()
Dim T$, S$, A$
S = "sllksdfj lsdkjf skldfj skldfj lk;asjdf lksjdf lsdkfjsdkflj "
T = TmpFt
WrtStr S, T
Debug.Assert SizFfn(T) = Len(S)
A = FfnBlk(T, 1)
Debug.Assert A = Left(S, 128)
End Sub

Function FnoBlk$(Fno%, IBlk)
Dim A As String * 128
Get #Fno, IBlk, A
FnoBlk = A
End Function

Function FfnBlk$(Ffn, IBlk)
Dim F%: F = FnoRnd(Ffn, 128)
FfnBlk = FnoBlk(F, IBlk)
Close #F
End Function
Sub ChktFxa(Ffn, Optional Fun$)
If Not IsFxa(Ffn) Then Thw Fun, "Given Ffn is not Fxa", "Ffn", Ffn
End Sub
Function IsFxa(Ffn) As Boolean
IsFxa = LCase(Ext(Ffn)) = FxaExt
End Function
Function IsFba(Ffn) As Boolean
IsFba = LCase(Ext(Ffn)) = FbaExt
End Function
Function IsPjf(Ffn) As Boolean
Select Case True
Case IsFxa(Ffn), IsFba(Ffn): IsPjf = True
End Select
End Function
Function IsFb(Ffn) As Boolean
Select Case LCase(Ext(Ffn))
Case FbExt, FbExt1: IsFb = True
End Select
End Function

Function IsFx(Ffn) As Boolean
Select Case LCase(Ext(Ffn))
Case ".xls", ".xlsm", ".xlsx": IsFx = True
End Select
End Function
Function FxAyzFfnAy(Ffny$()) As String()
Dim Ffn
For Each Ffn In Itr(Ffny)
    If IsFx(Ffn) Then PushI FxAyzFfnAy, Ffn
Next
End Function

Function FbAyzFfnAy(Ffny$()) As String()
Dim Ffn: For Each Ffn In Itr(Ffny)
    If IsFb(Ffn) Then PushI FbAyzFfnAy, Ffn
Next
End Function

Sub AsgExiMis(OExi$(), OMis$(), _
Ffny$())
Dim Ffn
Erase OExi
Erase OMis
For Each Ffn In Itr(Ffny)
    If HasFfn(Ffn) Then
        PushI OExi, Ffn
    Else
        PushI OMis, Ffn
    End If
Next
End Sub

Function HasFfn(Ffn) As Boolean
HasFfn = Fso.FileExists(Ffn)
End Function

Function NoFfn(Ffn) As Boolean
If Not HasFfn(Ffn) Then Debug.Print "NoFfn: " & Ffn: NoFfn = True
End Function

Function ExiFfnAet(Ffny$()) As Dictionary
Set ExiFfnAet = Aet(FfnAywExi(Ffny))
End Function

Function MisFfnAet(Ffny$()) As Dictionary
Set MisFfnAet = Aet(FfnAywMis(Ffny))
End Function

Function FfnAywExi(Ffny$()) As String()
Dim F: For Each F In Itr(Ffny)
    If HasFfn(F) Then PushI FfnAywExi, F
Next
End Function
Function FfnAywMis(Ffny$()) As String()
Dim F: For Each F In Itr(Ffny)
    If NoFfn(F) Then PushI FfnAywMis, F
Next
End Function

Sub ChkFfn(Ffn, Optional Fun$, Optional Kd$)
ChkFfnExist Ffn, Fun, Kd
End Sub

Sub ChkFfnExist(Ffn, Optional Fun$, Optional Kd$ = "File")
If NoFfn(Ffn) Then Thw Fun, "File not found", "File-Pth File-Name File-Kind", Pth(Ffn), Fn(Ffn), Kd
End Sub

Sub ChkFfnMis(Ffn, Fun$, Optional FilKind$)
If HasFfn(Ffn) Then Thw Fun, "File already exist", "File-Pth File-Name File-Kind", Pth(Ffn), Fn(Ffn), FilKind
End Sub

Function RplExt$(Ffn, NewExt)
RplExt = RmvExt(Ffn) & NewExt
End Function

Function DtezFfn(Ffn) As Date
If HasFfn(Ffn) Then DtezFfn = FileDateTime(Ffn)
End Function

Function SizFfn&(Ffn)
If NoFfn(Ffn) Then SizFfn = -1: Exit Function
SizFfn = FileLen(Ffn)
End Function

Function SiDotDTim$(Ffn)
If HasFfn(Ffn) Then SiDotDTim = TimStr(DtezFfn(Ffn)) & "." & SizFfn(Ffn)
End Function

Function TimStrzFfn$(Ffn)
TimStrzFfn = TimStr(DtezFfn(Ffn))
End Function

Function AddTimSfxzFfn$(Ffn)
AddTimSfxzFfn = AddFnSfx(Ffn, Format(Now, "(HHMMSS)"))
End Function
Function AddFnPfx$(A$, Pfx$)
AddFnPfx = Pth(A) & Pfx & Fn(A)
End Function

Function AddFnSfx$(Ffn, Sfx$)
AddFnSfx = RmvExt(Ffn) & Sfx & Ext(Ffn)
End Function


Function NxtNozFfn%(Ffn)
Dim A$: A = Right(RmvExt(Ffn), 5)
If FstChr(A) <> "(" Then Exit Function
If LasChr(A) <> ")" Then Exit Function
Dim M$: M = Mid(A, 2, 3)
If Not IsDigStr(M) Then Exit Function
NxtNozFfn = M
End Function
Function RmvNxtNo$(Ffn)
If IsNxtFfn(Ffn) Then
    Dim A$: A = RmvExt(Ffn)
    RmvNxtNo = RmvLasNChr(A, 5) & Ext(Ffn)
Else
    RmvNxtNo = Ffn
End If
End Function

Function InstFfn$(Ffn)
InstFfn = Pthi(Pth(Ffn)) & Fn(Ffn)
End Function

Function InstFdr$(Fdr)
InstFdr = AddFdrEns(TmpFdr(Fdr), NowStr)
End Function

Function CrtPthzInst$(Pth)
CrtPthzInst = Pthi(Pth)
End Function

Function IsInstFfn(Ffn) As Boolean
IsInstFfn = IsInstFdr(FdrzFfn(Ffn))
End Function

Function IsInstFdr(Fdr$) As Boolean
IsInstFdr = IsTimStr(Fdr)
End Function

Function FfnzPthFn$(Pth, Fn$)
FfnzPthFn = Ffn(Pth, Fn)
End Function

Function Ffn$(Pth, Fn$)
Ffn = EnsPthSfx(Pth) & Fn
End Function

Function HasExtSS(Ffn, ExtSS) As Boolean
Dim E$: E = Ext(Ffn)
Dim Sy$(): Sy = SyzSS(ExtSS)
HasExtSS = HasStrEle(Sy, E)
End Function

Sub OvrWrt(Ffn$, ShdOvrWrt As Boolean)
If ShdOvrWrt Then
    DltFfnIf Ffn
Else
    ChkNoFfn Ffn
End If
End Sub

Sub ChkHasFfn(Ffn)
If Not HasFfn(Ffn) Then
    Thw CSub, "File should exist", "Ffn", Ffn
End If
End Sub

Sub ChkNoFfn(Ffn)
If Not NoFfn(Ffn) Then
    Thw CSub, "File should not exist", "Ffn", Ffn
End If
End Sub

Function IsXlsx(Fn) As Boolean: IsXlsx = HasSfx(Fn, ".xlsx"): End Function
Function IsXlsFn(Fn) As Boolean:  IsXlsFn = HasSfx(Fn, ".xls"): End Function

Sub EnsFfnFm(Ffn$, FmFfn$)
If HasFfn(Ffn) Then Exit Sub
CpyFfn FmFfn, Ffn
End Sub
