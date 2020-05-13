Attribute VB_Name = "MxVbDtaS12Fmt"
Option Compare Text
Option Explicit
Const CNs$ = "S12"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbDtaS12Fmt."
Private Type Pm
    Tit As String
    Fmt As eTblFmt
    H1 As String
    H2 As String
    S1() As String
    S2() As String
    '---- Calc
    Top As String
    Sepln As String     'If NoRowSep, alway - will be replaced by ' ', otherwise, it will be HdrSep
    W() As Integer      '3 ele
    DtaQ() As String    '4 Ele
End Type
Private Type Bld
    Tit() As String ' Box
    Top As String   ' Top Ln Option
    Hdr As String   ' Hdr Ln
    HdrUL As String
    Middle() As String
End Type

'==Mov to other module
Function TopLn$(W%(), HasColSep As Boolean) ' #FmtTbl-Top-Ln# @W may have 0-value-ele.
Dim N%: N = Si(W)
Dim Q$(): Q = ColSepAy(W, HasColSep)
Dim Dr$()
Dim I: For Each I In W
    PushI Dr, String(I, "-")
Next
TopLn = FmtLn(Q, W, Dr)
End Function

Function ColSepAy(W%(), HasColSep As Boolean, Optional IsDta As Boolean) As String() ' #Column-Separator-Array# used in separating each column of formatting a table.  This will have 1 ele more than @W
If HasColSep Then
    ColSepAy = W1WiColSep(W, IsDta)
Else
    ColSepAy = W1NoColSep(W)
End If
End Function
Private Function W1WiColSep(W%(), IsDta As Boolean) As String()
Dim U%: U = UB(W)
Dim O$(): ReDim O(U + 1)
Dim C$: C = IIf(IsDta, " ", "-")
Dim J%
    O(0) = "|" & C
    For J = 1 To U - 1
        If W(J - 1) > 0 Then
            O(J) = C & "|" & C
        End If
    Next
    O(U) = C & "|"
W1WiColSep = O
End Function

Private Function W1NoColSep(W%()) As String()
Dim U%: U = UB(W)
Dim O$(): ReDim O(U + 1)
Dim J%: For J = 1 To U - 1
    If W(J - 1) > 0 Then
        O(J) = " "
    End If
Next
W1NoColSep = O
End Function
Function FmtLn$(Q$(), W%(), Dr, Optional AliRIxy)
Dim Ixy%(): If IsIntAy(AliRIxy) Then Ixy = AliRIxy
FmtLn = FmtLnzAliR(Q, W, Dr, Ixy)
End Function

Function FmtLnzAliR$(Q$(), W%(), Dr, AliRIxy%()) '#Fmt-Ln-With-AliR-Option# ret a formatted line of @Dr @Drasformatted line froy ..Pm
'@Q #QuoteStrAy#  QuoteStr between each @Dr.  Si = Si-Dr + 1
'@W #WdtAy#       Wdt of each ele of @Dr.     Si = Si-Dr
'@Dr #DtaRow#     value to be formatted.      Si > 0
'@AliRIxy #Ali-Right-Ix-Ay# which element in @Dr should be Align-Right, otherwise align-left
Dim O$(), A$(): A = W1Ali(W, Dr, AliRIxy)
PushI O, Q(0)
Dim J%: For J = 0 To UB(Q) - 1
    PushI O, A(J)
    PushI O, Q(J + 1)
Next
FmtLnzAliR = RTrim(Jn(O))
End Function

Private Function W1Ali(W%(), Dr, AliRIxy%()) As String()
Dim J%, Ixy%(): If IsIntAy(AliRIxy) Then Ixy = AliRIxy
Dim V: For Each V In Itr(Dr)
    If HasEle(AliRIxy, J) Then
        PushI W1Ali, AliR(V, W(J))
    Else
        PushI W1Ali, Ali(V, W(J))
    End If
    J = J + 1
Next
End Function

'**FmtS12y-2Lines
Private Sub FmtS12y2Ln__Tst()
End Sub
Function FmtS12y2Ln(A() As S12, Optional BegIx% = 1, Optional SndLnNSpc% = 4, Optional LnWdt% = 100) As String()
Dim Ix%: Ix = BegIx
Dim IxWdt%
Dim J%: For J = 0 To S12UB(A)
    PushIAy FmtS12y2Ln, W1OneS12(A(J), IxWdt, Ix, SndLnNSpc, LnWdt)
Next
End Function
Private Function W1OneS12(A As S12, IxWdt%, OIx%, SndLnTab%, LnWdt%) As String()
PushIAy W1OneS12, W1S1(A.S1, IxWdt, OIx)
Dim SndLnNSpc%
PushIAy W1OneS12, W1S2(A.S2, SndLnNSpc, LnWdt)
End Function
Private Function W1S1(S1$, IxWdt%, OIx%) As String()

End Function
Private Function W1S2(S1$, SndLnIndt%, LnWdt%) As String()
W1S2 = WrpLn(S1, LnWdt, SndLnIndt)
End Function

Private Sub FmtS12y__Tst()
Const ResPseg$ = "MxFmtS12\"
Dim A() As S12, H12$

'GoSub T0
'GoSub T1
GoSub T2
'GoSub T3
Exit Sub
T3:
    H12 = "AA BB"
    A = ResS12y(ResPseg & "Cas1\Inp_S12y.Txt")
    Ept = Resl(ResPseg & "Cas1\Ept.txt")
    GoTo Tst
T0:
    H12 = "AA BB"
    A = AddS12(S12("A", "B"), S12("AA", "B"))
    GoTo Tst
T1:
    H12 = "AA BB"
    A = SampS12y
    GoTo Tst
T2:
    H12 = "AA BB"
    A = SampS12y
    Brw FmtS12y(A, H12)
    Stop
    GoTo Tst
Tst:
    Act = FmtS12y(A, H12)
    C
    Return
End Sub

'==FmtS12y
Function FmtS12y(A() As S12, Optional H12$ = "S1 S2", Optional BegIx% = 1, Optional Fmt As eTblFmt, Optional Tit$) As String()
If S12Si(A) = 0 Then PushI FmtS12y, "(NoRec-S12y) N12=[" & H12 & "]"
Dim P As Pm: P = VVPm(A, H12, BegIx, Fmt, Tit)
With VVBld(P)
    FmtS12y = SyNB(.Tit, .Top, .Hdr, .HdrUL, .Middle, .Top)
End With
End Function

Private Function VVPm(A() As S12, H12$, BegIx%, Fmt As eTblFmt, Tit$) As Pm
Dim H As S12: H = BrkTRst(H12)
With VVPm
    .Fmt = Fmt
    .Tit = Tit
    .H1 = H.S1
    .H2 = H.S2
    .S1 = S1y(A)
    .S2 = S2y(A)

Dim HasIx As Boolean: HasIx = BegIx >= 0
Dim HasColSep As Boolean: HasColSep = Fmt = eBothSep Or Fmt = eColSep
ReDim .W(2)
    .W(0) = ZpmIxWdt(HasIx, UB(.S1))
    .W(1) = WdtzLinesy(AddSyStr(.S1, .H1))
    .W(2) = WdtzLinesy(AddSyStr(.S2, .H2))
    .DtaQ = ZpmDtaQ(HasColSep, HasIx)
    .Top = Sepln(.W, HasColSep)
    .Sepln = ZpmSepln(.Top, Fmt)
End With
End Function

Private Function VVBld(P As Pm) As Bld
Dim O As Bld
Dim ColSep As Boolean: ColSep = ZZColSep(P.Fmt)
With P
    O.Tit = Box(.Tit)
    O.Top = ZbldTop(.Fmt, .Top)
    O.Hdr = ZbldHdr(.W, .H1, .H2, ColSep)
    O.HdrUL = ZbldHdrUL(.H1, .H2, ColSep, .W)
    O.Middle = VVMid(P)
End With
VVBld = O
End Function

Private Function VVMid(P As Pm) As String()  ' Ret middle part of format A as eNoSep
Dim Ix&
With P
    Dim U&: U = UB(P.S1)
    Dim J&: For J = 0 To U
        PushIAy VVMid, ZmidOneMid(P.DtaQ, .W, .S1(J), .S2(J), .Sepln, Ix, J = U)
    Next
End With
End Function

Private Function ZpmDtaQ(ColSep As Boolean, HasIx As Boolean) As String() ' 3 ele
Dim A$, B$, C$, D$
    Select Case True
    Case ColSep And HasIx: A = "| ": B = " | ": C = " | ": D = " |"
    Case ColSep:           A = "| ": B = "":    C = " | ": D = " |"
    Case HasIx:            A = "":   B = " ":   C = " ":   D = ""
    Case Else:             A = "":   B = "":    C = " ":   D = ""
    End Select
ZpmDtaQ = Sy(A, B, C, D)
End Function
Private Function ZpmSepln$(Top$, Fmt As eTblFmt)
Select Case True
Case Fmt = eNoSep
Case Fmt = eBothSep Or Fmt = eRowSep: ZpmSepln = Top
Case Fmt = eColSep: ZpmSepln = Replace(Top, "-", " ")
Case Else: EnmEr CSub, "eTblFmtSS", eTblFmtSS, Fmt
End Select
End Function

Function ZpmIxWdt%(HasColSep As Boolean, U&)
If HasColSep Then ZpmIxWdt = Len(CStr(U + 1))
End Function

Private Function ZmidOneMid(Q$(), W%(), S1$, S2$, Sepln$, OIx&, IsLas As Boolean) As String()
Dim L1$(), L2$()
    L1 = SplitCrLf(S1)
    L2 = SplitCrLf(S2)
          ResiMax L1, L2
Dim IxStr$, IxSpc$
IxStr = ZmidIxStr(OIx, S1)
Dim AliRIxy%(0): AliRIxy(0) = 0 ' Ali first Ix col
Dim J&: For J = 0 To UB(L1)
    Dim Ix$: Ix = IIf(J = 0, IxStr, IxSpc)
    PushI ZmidOneMid, FmtLn(Q, W, Array(Ix, L1(J), L2(J)), AliRIxy)
Next
If Si(L1) > 1 Or IsLas Then
    PushI ZmidOneMid, Sepln
End If
End Function

Private Function ZmidIxStr$(OIx&, S1$)
Dim A$
If OIx >= 0 Then
    If IsHdrln(S1) Then
        A = OIx
    End If
End If
If A <> "" Then OIx = OIx + 1
ZmidIxStr = A
End Function

Private Function ZbldHdr$(W%(), H1$, H2$, HasColSep As Boolean) ' @W has 3 ele
Dim I$: I = String(W(0), "#")
Dim Q$(): Q = ColSepAy(W, HasColSep)
ZbldHdr = FmtLn(Q, W, Array(I, H1, H2))
End Function

Private Function ZbldTop$(F As eTblFmt, Sepln$)
If ZbldHasTop(F) Then ZbldTop = Sepln
End Function

Private Function ZbldHasTop(Fmt As eTblFmt) As Boolean
ZbldHasTop = Fmt = eRowSep Or Fmt = eBothSep
End Function

Private Function ZbldHdrUL$(H1$, H2$, ColSep As Boolean, W%())

End Function

Private Function W1HdrUL$(H1$, H2$, W%(), HasColSep As Boolean)
Dim A$, B$, C$
A = String(W(0), "=")
B = String(Len(H1), "=")
C = String(Len(H2), "=")
Dim Q$(): Q = ColSepAy(W, HasColSep)
W1HdrUL = FmtLn(Q, W, Array(A, B, C))
End Function

Private Function ZZRowSep(Fmt As eTblFmt) As Boolean: ZZRowSep = Fmt = eRowSep Or Fmt = eBothSep: End Function
Private Function ZZColSep(Fmt As eTblFmt) As Boolean: ZZColSep = Fmt = eColSep Or Fmt = eBothSep: End Function
Private Function ZZHasTop(Fmt As eTblFmt) As Boolean: ZZHasTop = ZZColSep(Fmt): End Function
