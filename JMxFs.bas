Attribute VB_Name = "JMxFs"
Option Compare Text
Const CMod$ = CLib & "JMxFs."
#If False Then
Option Explicit
Public Fso As New FileSystemObject
'---==================================================================================================== AA_Fs_Op
Private Function Msg$(FfnFm$, FfnTo$)
Dim F1$, F2$, P1$, P2$
F1 = Nam_FilNam(FfnFm)
F2 = Nam_FilNam(FfnFm)
P1 = Pth(FfnFm)
P2 = Pth(FfnTo)
Dim O$()
If F1 = F2 Then
    PushS O, vbTab & "File       : " & F1
Else
    PushS O, vbTab & "From File  : " & F1
    PushS O, vbTab & "To  File   : " & F2
End If
PushS O, vbTab & "From Folder: " & P1
PushS O, vbTab & "To   Folder: " & P2
If F1 = F2 Then
    PushS O, vbTab & FfnSiTimStr(FfnFm)
Else
    PushS O, vbTab & "From " & FfnSiTimStr(FfnFm)
    PushS O, vbTab & "To   " & FfnSiTimStr(FfnTo)
End If
PushS O, ""
Msg = vbCrLf & Join(O, vbCrLf)
End Function
Private Function FfnSiTimStr$(Ffn$)
Dim S$: S = Format(FileLen(Ffn), "##,###,###,###")
S = AlignR(S, 15)
FfnSiTimStr = "File Size/Time: [" & AlignR(Format(FileLen(Ffn), "##,###,###,###"), 15) & "] [" & Format(FileDateTime(Ffn), "YYYY-MM-DD HH:MM:SS") & "]"
End Function
Sub CpyFfnAy_IfDif(AyFfn$(), ToPth$)
Dim F: For Each F In Itr(AyFfn)
    CpyFfnIfDif CStr(F), ToPth & Nam_FilNam(CStr(F))
Next
End Sub
Sub DltFfn(Ffn$)
On Error GoTo X
Kill Ffn
Exit Sub
X: Err.Raise 1, , Err.Description & vbCrLf & vbCrLf & "Cannot delete file:" & Ffn2LinStr(Ffn)
End Sub
Sub DltFfnIf(Ffn$)
If Dir(Ffn) <> "" Then DltFfn Ffn
End Sub

Sub CpyFfn(Fm$, ToFil$)
On Error GoTo E
ChkPthExist Pth(ToFil)
DltFfnIf ToFil
FileCopy Fm$, ToFil
Exit Sub
E: MsgBox "Error in copying: " & Err.Description & vbCrLf & "From: " & vbCrLf & Fm & vbCrLf & "To:" & ToFil
End Sub
'---==================================================================================================== AA_Fs_Chk
Function Pth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn, P)
End Function
Function Fn$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Fn = Ffn: Exit Function
Fn = Mid(Ffn, P + 1)
End Function
Function Ffn2LinStr$(Ffn)
Ffn2LinStr = "File[" & Fn(Ffn) & "]" & vbCrLf & "Folder[" & Pth(Ffn) & "]"
End Function

Sub ChkFfnExist(Ffn$, Optional Kd$ = "File")
If Dir(Ffn) = "" Then Thw Kd & " not found:" & vbCrLf & vbCrLf & Ffn2LinStr(Ffn)
End Sub

Sub ChkPthExist(Pth$)
If Dir(Pth, vbDirectory) = "" Then Raise "Folder not found: [" & Pth & "]"
End Sub


Function IsSamTimSi(Ffn1$, Ffn2$) As Boolean
If FileLen(Ffn1) <> FileLen(Ffn2) Then Exit Function
If FileDateTime(Ffn1) <> FileDateTime(Ffn2) Then Exit Function
IsSamTimSi = True
End Function

#End If
