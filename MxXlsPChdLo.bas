Attribute VB_Name = "MxXlsPChdLo"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsPChdLo."
Sub PutPChdLo(SrcLo As ListObject, Gpcc$, At As Range)
Dim GpFny$(): GpFny = SyzSS(Gpcc)
Dim SrcFny$(): SrcFny = FnyzLo(SrcLo)
PutSq ParSq(SrcLo, GpFny), At
PutSq ChdSq(SrcLo, GpFny), ChdAt(At, SrcFny, GpFny)
AddWsSrc WszLo(SrcLo), PChdLoSrcl
End Sub

Private Function ParSq(SrcLo As ListObject, GpFny$()) As Variant()

End Function
Private Function ChdSq(SrcLo As ListObject, GpFny$()) As Variant()

End Function
Private Function ChdAt(At As Range, SrcFny$(), GpFny$()) As Range

End Function
Private Function PChdLoSrcl$()

End Function
