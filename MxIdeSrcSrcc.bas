Attribute VB_Name = "MxIdeSrcSrcc"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcSrcc."

Function ClnSrcP() As String(): ClnSrcP = ClnSrczP(CPj): End Function ':Src #Clean-Src# Src without Blnk/Rmk line

Function SrcHasSngDblQ(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If HasSngDblQ(L) Then
        PushI SrcHasSngDblQ, L
    End If
Next
End Function

Function HasSngDblQ(S) As Boolean
If HasSngQ(S) Then
    If HasDblQ(S) Then
        HasSngDblQ = True
    End If
End If
End Function

Function ClnSrczM(M As CodeModule) As String()
ClnSrczM = ClnSrc(Src(M))
End Function

Function ClnSrczP(P As VBProject) As String(): ClnSrczP = ClnSrc(SrczP(P)): End Function

Function ClnSrcs(ClnSrc$()) As String() ':Src #Clean-Src-Without-VbStrCxt# ! All string-cxt quoted by DblQ are removed
Dim L: For Each L In Itr(ClnSrc)
    PushI ClnSrcs, RmvVbStr(L)
Next
End Function

Function RmvVbStr$(NoVrmk)
Dim O$: O = RplDblSpc(NoVrmk)
Dim J%
X:
    LoopTooMuch CSub, J
    Dim P1&: P1 = InStr(O, vbDblQ): If P1 = 0 Then RmvVbStr = NoVrmk: Exit Function
    Dim P2&: P2 = InStr(P1 + 1, O, vbDblQ): If P2 = 0 Then Stop
    O = Left(O, P1 - 1) & Mid(O, P2 + 1)
    GoTo X
End Function

Function ClnSrc(Src$()) As String() ' :Src #Clean-Src# ! all empty-Ln and rmk-Ln are removed.
Dim L: For Each L In Itr(Src)
    If Not IsRmkOrBlnk(L) Then PushI ClnSrc, L
Next
End Function

Function IsCdLn(L) As Boolean: IsCdLn = Not IsRmkOrBlnk(L): End Function

Function IsLnNonOpt(Ln) As Boolean
If Not IsCdLn(Ln) Then Exit Function
If HasPfx(Ln, "Option") Then Exit Function
IsLnNonOpt = True
End Function
