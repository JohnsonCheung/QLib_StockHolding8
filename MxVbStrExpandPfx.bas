Attribute VB_Name = "MxVbStrExpandPfx"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbStrExpandPfx."

Function ExpandPfxSS$(Pfx$, SS$)
Dim O$()
Dim I: For Each I In Split(SS)
    PushS O, Pfx & I
Next
ExpandPfxSS = Join(O, " ")
End Function
Function ExpandPfxNN$(Pfx$, Fst%, Las%, Optional Fmt$, Optional Sep$ = " ")
ExpandPfxNN = Join(ExpandPfxNNAy(Pfx, Fst, Las, Fmt), Sep)
End Function
Function ExpandPfxNNAy(Pfx$, Fst%, Las%, Optional Fmt$) As String()
Dim I%: For I = Fst To Las
    PushS ExpandPfxNNAy, Pfx & Format(I, Fmt)
Next
End Function
