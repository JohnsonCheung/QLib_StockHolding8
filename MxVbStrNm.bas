Attribute VB_Name = "MxVbStrNm"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrNm."

Function IsNm(S) As Boolean
If S = "" Then Exit Function
If Not IsLetter(FstChr(S)) Then Exit Function
Dim L&: L = Len(S)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(S, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsNmChr(C$) As Boolean
IsNmChr = True
If IsLetter(C) Then Exit Function
If C = "_" Then Exit Function
If IsDigit(C) Then Exit Function
IsNmChr = False
End Function

Function IsDotNmChr(A$) As Boolean
If IsNmChr(A) Then IsDotNmChr = True: Exit Function
IsDotNmChr = A = "."
End Function

Function HitRx(S, Rx As RegExp) As Boolean
If S = "" Then Exit Function
If IsNothing(Rx) Then Exit Function
HitRx = Rx.Test(S)
End Function

Function NxtSeqNm$(SeqNm$, Optional NDig% = 3) _
':SeqNm: :Nm ! #Seq-Nm#  it can be XXX or XXX_nn where nn can be 1, 2 or 3 digits
'   If XXX, return XXX_001   '<-- # of zero depends on NDig
'   If XXX_nn, return XXX_mm '<-- mm is nn+1, # of digit of nn and mm depends on NDig
If Not IsBet(NDig, 1, 7) Then PmEr CSub, "Should between 1 and 7", "NDig", NDig
Dim HasSeq As Boolean: Stop
    Dim R$: R = Right(SeqNm, NDig + 1)
    Select Case True
    Case Len(R) = NDig + 1
    Case FstChr(R) = "_"
    Case Not IsNumeric(RmvFstChr(R))
    Case Else
        HasSeq = True
    End Select
If Not HasSeq Then
    NxtSeqNm = SeqNm & "_" & Pad0(1, NDig)
    Exit Function
End If
Dim Nm$, Sfx$
    Dim NmLen%, Seq&: Stop
    Nm = Left(SeqNm, NmLen)
    Sfx = Pad0(Seq + 1, NDig)
NxtSeqNm = Nm & "_" & Sfx
End Function

Function TakDotNm$(S)
Dim L%: L = Len(S): If L = 0 Then Exit Function
If Not IsLetter(FstChr(S)) Then Exit Function
Dim J%: For J = 2 To L
    If Not IsDotNmChr(Mid(S, J, 1)) Then
        TakDotNm = Left(S, J - 1)
        Exit Function
    End If
Next
TakDotNm = S
End Function

Function Ny(Lines) As String()
Ny = Idry(Lines)
End Function

Function NyzSy(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    PushI NyzSy, TakNm(S)
Next
End Function

Function AftNm$(S)
AftNm = Mid(S, Len(TakNm(S)) + 1)
End Function

Function Nm1$(S)
If S = "" Then Exit Function
If Not IsLetter(FstChr(S)) Then Exit Function
Dim J%: For J = 2 To Len(S)
    If Not IsNmChr(Mid(S, J, 1)) Then
        Nm1 = Left(S, J - 1)
        Exit Function
    End If
Next
Nm1 = S
End Function

Function TakNm$(S)
TakNm = Nm1(S)
End Function


Sub ChkNy(Ny$(), Fun$)
Dim N: For Each N In Itr(Ny)
    If Not IsNm(N) Then Thw Fun, "Ele of Sy is not nm", "Not-nm-Ele Sy", N, Sy
Next
End Sub

Function ShfNm$(OLin$)
Dim O$: O = TakNm(OLin): If O = "" Then Exit Function
ShfNm = O
OLin = RmvPfx(OLin, O)
End Function


Function ShfDotNm$(OLin$)
OLin = LTrim(OLin)
Dim O$: O = TakDotNm(OLin): If O = "" Then Exit Function
ShfDotNm = O
OLin = RmvPfx(OLin, O)
End Function
