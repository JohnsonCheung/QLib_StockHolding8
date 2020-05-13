Attribute VB_Name = "MxVbStrSfxEns"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrSfxEns."

Function EnsAyDotSfx(Ay) As String()
EnsAyDotSfx = EnsAySfx(Ay, ".")
End Function

Function EnsAySfx(Ay, Sfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushI EnsAySfx, EnsSfx(I, Sfx)
Next
End Function

Function EnsSfx(S, Sfx)
If HasSfx(S, Sfx) Then
    EnsSfx = S
Else
    EnsSfx = S & Sfx
End If
End Function

Function EnsSfxDot$(S)
EnsSfxDot = EnsSfx(S, ".")
End Function

Function EnsSfxSemi$(S)
EnsSfxSemi = EnsSfx(S, ";")
End Function
