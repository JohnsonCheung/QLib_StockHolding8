Attribute VB_Name = "MxVbFsPth"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsPth."

':Pseg: :S #Pth-Segment# ! 0 or more :Fdr separated by :PthSep
':Fdr:  :S #Folder#      ! An directory entry in file-system
Function IsPseg(Pseg$) As Boolean
Select Case True
Case FstChr(Pseg) = "\"
Case LasChr(Pseg) = "\"
Case Else: IsPseg = True
End Select
End Function

Function AddFdr$(Pth, Fdr)
AddFdr = EnsPthSfx(Pth) & AddNBAp(Fdr, "\")
End Function

Function AddPseg$(Pth, Pseg)
'Ret : :Pth
AddPseg = EnsPthSfx(EnsPthSfx(Pth) & Pseg)
End Function

Function AddFdrEns$(Pth, Fdr)
AddFdrEns = EnsPth(AddFdr(Pth, Fdr))
End Function

Function AddFdrApEns$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
Dim O$: O = AddFdrAv(Pth, Av)
EnsAllFdr O
AddFdrApEns = O
End Function

Function AddFdrAv$(Pth, FdrAv())
Dim O$: O = Pth
Dim I, Fdr$
For Each I In FdrAv
    Fdr = I
    O = AddFdr(O, Fdr)
Next
AddFdrAv = O
End Function

Function AddFdrAp$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
AddFdrAp = AddFdrAv(Pth, Av)
End Function

Function IsEmpPth(Pth) As Boolean
ChkPthExist Pth, CSub
If AnyFil(Pth) Then Exit Function
If HasSubFdr(Pth) Then Exit Function
IsEmpPth = True
End Function

Function AddFdrPfx$(Pth, Pfx)
With Brk2Rev(RmvPthSfx(Pth), PthSep, NoTrim:=True)
    AddFdrPfx = .S1 & PthSep & Pfx & .S2 & PthSep
End With
End Function

Function HitFilAtr(A As VbFileAttribute, Wh As VbFileAttribute) As Boolean
HitFilAtr = True
End Function

Function FdrzFfn$(Ffn)
FdrzFfn = Fdr(Pth(Ffn))
End Function

Function Fdr$(Pth)
Fdr = AftRev(RmvPthSfx(Pth), PthSep)
End Function

Sub ChktProperFdrNm(Fdr$)
Const CSub$ = CMod & "ChktProperFdrNm"
Const C$ = "\/:<>"
If HasChrList(Fdr, C) Then Thw CSub, "Fdr cannot has these char " & C, "Fdr Char", Fdr, C
End Sub

Function RmvFdr$(Pth)
RmvFdr = BefRev(RmvPthSfx(Pth), PthSep) & PthSep
End Function

Function ParPth$(Pth) ' Return the ParPth of given Pth
ParPth = RmvFdr(Pth)
End Function

Function ParFdr$(Pth)
ParFdr = Fdr(ParPth(Pth))
End Function

Function UpNFdr$(Pth, UpN%)
Dim O$: O = Pth
Dim J%: For J = 1 To UpN
    O = ParPth(O)
Next
UpNFdr = O
End Function

Function PthInst$(Pth)
Dim P$: P = EnsPthSfx(Pth)
PthInst = EnsPth(P & NxtFdr(P) & "\")
End Function

Function EnsPth$(Pth)
Dim P$: P = EnsPthSfx(Pth)
If NoPth(P) Then MkDir RmvLasChr(P)
EnsPth = P
End Function

Function EnsFfnAllFdr$(Ffn)
EnsAllFdr Pth(Ffn)
EnsFfnAllFdr = Ffn
End Function

Function EnsAllFdr$(Pth)
'Ret :Pth and ens each :Pseg. @@
Dim J%, O$, Ay$()
Ay = Split(RmvSfx(Pth, PthSep), PthSep)
O = Ay(0)
For J = 1 To UBound(Ay)
    O = O & PthSep & Ay(J)
    EnsPth O
Next
EnsAllFdr = EnsPthSfx(Pth)
End Function

Function HasPth(Pth) As Boolean
HasPth = Fso.FolderExists(Pth)
End Function

Function NoPth(Pth) As Boolean
If Not HasPth(Pth) Then Debug.Print "NoPth: "; Pth: NoPth = True
End Function

Function HasFdr(Pth, Fdr$) As Boolean
HasFdr = HasEle(FdrAy(Pth), Fdr)
End Function

Sub ChkPthExist(Pth, Fun$)
If NoPth(Pth) Then Thw Fun, "Pth not exist", "Pth", Pth
End Sub

Function AnyFil(Pth) As Boolean
AnyFil = Dir(Pth) <> ""
End Function

Function HasSubFdr(Pth) As Boolean
HasSubFdr = Fso.GetFolder(Pth).SubFolders.Count > 0
End Function

Function FdrAyzIsInst(Pth) As String()
Dim I, Fdr$
For Each I In Itr(FdrAy(Pth))
    Fdr = I
    If IsInstNm(Fdr) Then PushI FdrAyzIsInst, Fdr
Next
End Function

Function FdrAyC(Optional Spec$ = "*.*") As String()
FdrAyC = FdrAy(Cd, Spec)
End Function
Function FdrAy(Pth, Optional Spec$ = "*.*") As String()
Dim P$: P = EnsPthSfx(Pth)
Dim E: For Each E In Itr(EntAy(P, Spec))
    If (GetAttr(P & E) And VbFileAttribute.vbDirectory) <> 0 Then
        PushI FdrAy, E    '<====
    End If
Next
End Function

Function EntAyC(Optional Spec$ = "*.*") As String()
EntAyC = EntAy(Cd, Spec)
End Function

Function EntAy(Pth, Optional Spec$ = "*.*") As String()
Const CSub$ = CMod & "EntAy"
If NoPth(Pth) Then Exit Function
Dim A$: A$ = Dir(EnsPthSfx(Pth) & Spec, vbDirectory)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If InStr(A, "?") > 0 Then
        Inf CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec", A, Pth, Spec
        GoTo X
    End If
    PushI EntAy, A
X:
    A = Dir
Wend
End Function
Function IsInstNm(Nm) As Boolean
If FstChr(Nm) <> "N" Then Exit Function      'FstChr = N
If Len(Nm) <> 16 Then Exit Function          'Len    =16
If Not IsYYYYMMDD(Mid(Nm, 2, 8)) Then Exit Function 'NYYYYMMDD_HHMMDD
If Mid(Nm, 10, 1) <> "_" Then Exit Function
If Not IsHHMMDD(Right(Nm, 6)) Then Exit Function
IsInstNm = True
End Function

Function FfnItr(Pth)
Asg Itr(Ffny(Pth)), FfnItr
End Function

Function SubPthAy(Pth) As String()
SubPthAy = AmAddPfxSfx(FdrAy(Pth), EnsPthSfx(Pth), PthSep)
End Function
Sub ChgCd(Pth)
ChkPthExist Pth, "ChgCd"
ChDir Pth
If Not HasPth(Pth) Then Thw CSub, "Pt"
End Sub
Function Cd$()
Cd = CurDir & "\"
End Function

Function CSubPthAy() As String()
CSubPthAy = SubPthAy(Cd)
End Function

Sub AsgEnt(OFdrAy$(), OFnAy$(), Pth)
Erase OFdrAy
Erase OFnAy
Dim A$, P$
P = EnsPthSfx(Pth)
A = Dir(Pth, vbDirectory)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If HasPth(P & A) Then
        PushI OFdrAy, A
    Else
        PushI OFnAy, A
    End If
    A = Dir
X:
Wend
End Sub

Function FnnAy(Pth, Optional Spec$ = "*.*") As String()
Dim I: For Each I In FnAy(Pth, Spec)
    PushI FnnAy, RmvExt(I)
Next
End Function

Function FnAyzFfnAy(Ffny$()) As String()
Dim I, Ffn$
For Each I In Itr(Ffny)
    Ffn = I
    PushI FnAyzFfnAy, Fn(Ffn)
Next
End Function

Function FnAy(Pth, Optional Spec$ = "*.*") As String()
Dim O$()
Dim M$: M = Dir(EnsPthSfx(Pth) & Spec)
While M <> ""
   PushI FnAy, M
   M = Dir
Wend
End Function

Function Fxy(Pth) As String()
Fxy = Ffny(Pth, "*.xls*")
End Function

Function Ffny(Pth, Optional Spec$ = "*.*") As String()
Ffny = AmAddPfx(FnAy(Pth, Spec), EnsPthSfx(Pth))
End Function

Private Sub SubPthAy__Tst()
Dim Pth
Pth = "C:\Users\user\AppData\Local\Temp\"
Ept = Sy()
GoSub Tst
Exit Sub
Tst:
    Act = SubPthAy(Pth)
    Brw Act
    Return
End Sub

Private Sub Fxy__Tst()
Dim A$()
A = Fxy(CurDir)
DmpAy A
End Sub

Function HasPthSfx(Pth) As Boolean
HasPthSfx = LasChr(Pth) = PthSep
End Function

Function EnsPthSfx$(Pth)
If Pth = "" Then Exit Function
If HasPthSfx(Pth) Then
    EnsPthSfx = Pth
Else
    EnsPthSfx = Pth & PthSep
End If
End Function

Function RmvPthSfx$(Pth)
RmvPthSfx = RmvSfx(Pth, PthSep)
End Function

Function HasSiblingFdr(Pth, Fdr$) As Boolean
HasSiblingFdr = HasFdr(ParPth(Pth), Fdr)
End Function

Function SiblingPth$(Pth, SiblingFdr$)
SiblingPth = AddFdrEns(ParPth(Pth), SiblingFdr)
End Function

Function Pthi$(Pth)
':Pthi: :Pth ! An Pthi is a SubPth of a given Pth. with Fdr = NowStr()
Pthi = AddFdrEns(Pth, NowStr)
End Function
