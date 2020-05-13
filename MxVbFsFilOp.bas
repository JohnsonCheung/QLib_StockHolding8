Attribute VB_Name = "MxVbFsFilOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsFilOp."
Sub RplFfn(Ffn, ByFfn$)
BkuFfn Ffn
If DltFfnDone(Ffn) Then
    Name Ffn As ByFfn
End If
End Sub
Sub CpyPthzClr(FmPth$, ToPth$)
ChkPthExist ToPth, CSub
DltAllPthFil ToPth
Dim Ffn$, I
For Each I In Ffny(FmPth)
    Ffn = I
    CpyFfnToPth Ffn, ToPth
Next
End Sub

Sub CpyFfnUp(Ffn)
CpyFfnToPth Ffn, ParPth(Ffn)
End Sub

Sub CpyFfnAyzToNxt(Ffny$())
Dim I, Ffn$
For Each I In Itr(Ffny)
    Ffn = I
    CpyFfnzToNxt Ffn
Next
End Sub

Function CpyFfnzToNxt$(Ffn)
Dim O$
O = NxtFfnzAva(Ffn)
CpyFfn Ffn, O
CpyFfnzToNxt = O
End Function

Sub CpyFfnToPthIfDif(Ffn, ToPth$, Optional B As eFilCprMth = eFilCprMth.eCprEachbyt)
Dim Fn$: Fn = FnzFfn(Ffn)
Dim ToFfn$: ToFfn = FfnzPthFn(ToPth, Fn)
If IsEqFfn(Ffn, ToFfn, B) Then Exit Sub
CpyFfnToPth Ffn, ToPth, OvrWrt:=True
End Sub

Sub CpyFfnAyzIfDif(Ffny$(), ToPth$, Optional B As eFilCprMth = eFilCprMth.eCprEachbyt)
Dim I
For Each I In Ffny
    CpyFfnIfDif CStr(I), ToPth, B
Next
End Sub

Sub CpyFfn(Ffn, ToFfn$, Optional OvrWrt As Boolean)
Fso.GetFile(Ffn).Copy ToFfn, OvrWrt
End Sub

Function CpyFfnAy$(Ffny$(), ToPth$, Optional OvrWrt As Boolean)
Dim Ffn$, I, P$, O$
P = EnsPthSfx(ToPth)
For Each I In Ffny
    O = P & Fn(Ffn)
    CpyFfn Ffn, O, OvrWrt
Next
End Function

Sub CpyFfnToPth(Ffn, ToPth$, Optional OvrWrt As Boolean)
CpyFfn Ffn, FfnzPthFn(ToPth, Fn(Ffn)), OvrWrt
End Sub

Sub CpyFfnIfDif(Ffn, ToFfn$, Optional M As eFilCprMth)
Const CSub$ = CMod & "CpyFfnzIfDif"
If IsEqFfn(Ffn, ToFfn, M) Then
    Dim Msg$: Msg = FmtQQ("? file", IIf(M = eCprEachbyt, "EachByt", "SamTimSi"))
    D FmtFmsgNap(CSub, Msg, "FmFfn ToFfn", Ffn, ToFfn)
    Exit Sub
End If
CpyFfn Ffn, ToFfn, OvrWrt:=True
D FmtFmsgNap(CSub, "File copied", "FmFfn ToFfn", Ffn, ToFfn)
End Sub

Sub DltFfnAyIf(Ffny$())
Dim Ffn: For Each Ffn In Itr(Ffny)
    DltFfnIf Ffn
Next
End Sub

Sub DltFfn(Ffn)
Const CSub$ = CMod & "DltFfn"
On Error GoTo X
Kill Ffn
Exit Sub
X:
Thw CSub, "Cannot kill", "Ffn Er", Ffn, Err.Description
End Sub

Sub DltFfnIf(Ffn)
If HasFfn(Ffn) Then DltFfn Ffn
End Sub

Function DltFfnIfPrompt(Ffn, Msg$) As Boolean 'Return true if error
If NoFfn(Ffn) Then Exit Function
On Error GoTo X
Kill Ffn
Exit Function
X:
MsgBox "File [" & Ffn & "] cannot be deleted, " & vbCrLf & Msg
DltFfnIfPrompt = True
End Function

Function DltFfnDone(Ffn) As Boolean
On Error GoTo X
Kill Ffn
DltFfnDone = True
Exit Function
X:
End Function


Sub MovFilUp(Pth)
Dim I, Tar$
Tar$ = ParPth(Pth)
For Each I In Itr(FnAy(Pth))
    MovFfn CStr(I), Tar
Next
End Sub


Sub MovFfn(Ffn, ToPth$)
Fso.MoveFile Ffn, ToPth
End Sub
