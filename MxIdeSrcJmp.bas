Attribute VB_Name = "MxIdeSrcJmp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcJmp."
Private Type MdnLcc
    Mdn As String
    Lno As Long
    C1 As Integer
    C2 As Integer
End Type

'**Jmp
Sub Jmp__Tst(): Jmp "MxGit:1": End Sub
Sub JmpCmpn(Cmpn$)
Dim C As VBIde.CodePane: Set C = PnezCmpn(Cmpn)
If IsNothing(C) Then Debug.Print "No such WinOfCmpNm": Exit Sub
C.Show
End Sub
Sub JmpMd(M As CodeModule): M.CodePane.Show: End Sub
Sub JmpMdLno(M As CodeModule, Lno&)
JmpMd M
JmpLno Lno
End Sub

Function MdnLcczS(MdnLccStr$) As MdnLcc
Dim A$(): A = SplitColon(MdnLccStr)
With MdnLcczS
    .Mdn = Shf(A)
    .Lno = Shf(A)
    .C1 = Shf(A)
    .C2 = Shf(A)
End With
End Function
Sub Jmp(MdnLccStr$) 'MdnLccStr: #Mdn-Lno-C1-C2-Str# fmt/Mdn:Lno:C1:C2/ where :C1:C2 is optional
With MdnLcczS(MdnLccStr)
If .Mdn <> "" Then JmpMdn .Mdn
If .Lno > 0 Then
    JmpLno .Lno
    If .C1 > 0 Then CPne.SetSelection .Lno, .C1, .Lno, .C2
End If
End With
End Sub

Sub JmpRCC(R&, C1%, C2%)
CPne.SetSelection R, C1, R, C2
End Sub

Sub JmpMdn(Mdn)
ClsAllWin
JmpMd Md(Mdn)
End Sub

Sub JmpLno(Lno&)
Dim C2%: C2 = Len(CMd.Lines(Lno, 1)) + 1
JmpLcc Lno, 1, C2
End Sub

Sub JmpLcc(Lno&, C1%, C2%)
Dim L1&: L1 = Lno - 6: If L1 <= 0 Then L1 = 1
With CPne
    .TopLine = L1
    .SetSelection Lno, C1, Lno, C2
End With
End Sub

Sub JmpRRCC(A As RRCC)
Dim L&, C1%, C2%
With CPne
    If C1 = 0 Or C2 = 0 Then
        C1 = 1
        C2 = Len(.CodeModule.Lines(L, 1)) + 1
    End If
    .TopLine = L
    .SetSelection L, C1, L, C2
End With
'SendKeys "^{F4}"
End Sub

Sub JmpMth(Mthn)
Dim M As CodeModule: Set M = MdzMthn(CPj, Mthn)
JmpMd M
JmpLno Mthlno(M, Mthn)
End Sub

Sub JmpMdMth(M As CodeModule, Mthn)
JmpMd M
JmpMth Mthn
End Sub

Sub JmpPj(P As VBProject)
ClsAllWin
Dim M As CodeModule
Set M = FstMd(P)
If IsNothing(M) Then Exit Sub
JmpMd M
TileV
DoEvents
End Sub

Sub JmpMdRRCC(M As CodeModule, R As RRCC)
JmpMd M
JmpRRCC R
End Sub

Sub JmpMdnn(Mdnn$)
ClsAllWin
Dim M: For Each M In Itr(SyzSS(Mdnn))
    VisMd Md(M)
Next
TileV
End Sub

Sub VisMd(M As CodeModule)
M.CodePane.Window.Visible = True
End Sub
