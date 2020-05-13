Attribute VB_Name = "MxIdeSrcDclUdtOp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclUdtOp."

Sub InsUdt(M As CodeModule, UdtCdl$, AftUdtn$)
If UdtCdl = "" Then Exit Sub
M.InsertLines LnoAftUdtn(M, AftUdtn), UdtCdl
End Sub

Sub DltUdt__Tst()
DltUdt CMd, "AA"
End Sub

Sub DltUdt(M As CodeModule, Udtn$) ' Delete Udt if exist else do nothing
Dim A As Bei: A = UdtBei(Dcl(M), Udtn): If IsEmpBei(A) Then Exit Sub
DltCdlzLcnt M, LcntzBei(A)
End Sub

Sub EnsUdt(M As CodeModule, UdtCdl$, AftUdtn$) '@M should have this @UdtCdl, otherwise dlt @Udtn and ins @UdtCdl to @M after @AftUdtn
If HasDclCdl(M, UdtCdl) Then Exit Sub
DltUdt M, Udtn(UdtCdl)
InsUdt M, UdtCdl, AftUdtn
End Sub

Private Function LnoAftUdtn%(M As CodeModule, Udtn$)
Dim A As Bei: A = UdtBei(Dcl(M), Udtn)
If A.Eix >= 0 Then LnoAftUdtn = A.Eix + 2: Exit Function
Thw CSub, "Udtn not found in Md", "Udtn Mdn", Udtn, Mdn(M)
'LnoAftUdtn = LasDclLno(M) + 1
End Function

Sub LasDclLno__Tst()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    Dim M As CodeModule: Set M = C.CodeModule
    Dim N%: N = LasDclLno(M)
    PushI O, JnTabAp(Mdn(M), N, M.Lines(N, 1))
    PushI O, JnTabAp(Mdn(M), N + 1, M.Lines(N + 1, 1))
Next
VcAy O
End Sub

Private Function LasDclLno%(M As CodeModule) ' This must be a code line
Dim N%: N = M.CountOfDeclarationLines
LasDclLno = N - NRmkBlnkLnAbove(M, M.CountOfDeclarationLines)
End Function
