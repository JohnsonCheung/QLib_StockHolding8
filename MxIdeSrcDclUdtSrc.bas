Attribute VB_Name = "MxIdeSrcDclUdtSrc"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclUdtSrc."

Private Sub UdtSrcy__Tst():  VcLyAy UdtSrcy(DclP):  End Sub
Private Sub UdtStmty__Tst(): VcLyAy UdtStmty(DclP): End Sub
Function UdtSrcl$(Dcl$(), Udtn$):          UdtSrcl = JnCrLf(UdtSrc(Dcl, Udtn)):     End Function
Function UdtSrc(Dcl$(), Udtn$) As String(): UdtSrc = AwBei(Dcl, UdtBei(Dcl, Udtn)): End Function
Function UdtSrcy(Dcl$()) As Variant()
Dim Bei() As Bei: Bei = UdtBeiy(Dcl)
Dim J%: For J = 0 To BeiUB(Bei)
    PushI UdtSrcy, AwBei(Dcl, Bei(J))
Next
End Function
Private Function UdtStmty(Dcl$()) As Variant(): UdtStmty = Stmty(UdtSrcy(Dcl)): End Function

