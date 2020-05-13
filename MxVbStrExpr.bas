Attribute VB_Name = "MxVbStrExpr"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CNs$ = "Lang.Vb"
Const CMod$ = CLib & "MxVbStrExpr."
Private Type LinRslt
    EprLin As String
    OvrFlwTerm As String
    S As String
End Type
Private Type Term
    EprTerm As String
    S As String
End Type

Function DqStrEpr(S, Optional MaxCdLinWdt% = 200) As String()
Dim L, Ay$(): Ay = SplitCrLf(S)
Dim J&, Fst As Boolean
Dim O$()
Fst = True
For Each L In Itr(Ay)
    If Fst Then
        Fst = False
    Else
        X J & ":" & Len(L) & ":" & L
    End If
    J = J + 1
    PushIAy DqStrEpr, DqStrEprLin(L, MaxCdLinWdt)
Next
DqStrEpr = O
End Function

Function DqStrEprLin(Ln, W%) As String()
Const CSub$ = CMod & "DqStrEprLin"
Dim J%
Dim S$: S = Ln
Dim CurLen&
Dim LasLen&: LasLen = Len(S)
Dim OvrFlwTerm$
While LasLen > 0
    DoEvents
    LoopTooMuch CSub, J
    Stop
    If J > 10 Then Stop
    With ShfLn(S, OvrFlwTerm, W)
        If .EprLin = "" Then Exit Function
        PushI DqStrEprLin, .EprLin
        S = .S
        OvrFlwTerm = .OvrFlwTerm
    End With
    CurLen = Len(S)
    If CurLen >= LasLen Then ThwNever CSub, "Str is not shifted by ShfLn"
    LasLen = CurLen
Wend
End Function

Function ShfLn(Str$, OvrFlwTerm$, W%) As LinRslt
Dim T$, OEprTermy, TotW&
If OvrFlwTerm <> "" Then
    PushI OEprTermy, OvrFlwTerm
    TotW = Len(OvrFlwTerm) + 3
End If
Dim S$: S = Str
Dim J&, OStr$, OEprTerm$

X:
ShfLn = LinRslt(EprLin:=Jn(OEprTermy, " & "), OvrFlwTerm:=OvrFlwTerm, S:=OStr)
End Function
Private Sub ShfTermzPrintable__Tst()
Dim S$: S = PjfStrP
Dim Las&, Cur&, O$()
Las = Len(S)
While Len(S) > 0
    PushI O, ShfTermzPrintable(S)
    Cur = Len(S)
    If Cur >= Las Then Stop
    Las = Cur
Wend
MsgBox Si(O)
Stop
Brw O
End Sub
Function ShfTermzPrintable$(OStr$)
If OStr = "" Then Exit Function
Dim IsPrintable As Boolean
Dim J&
IsPrintable = IsAscPrintable(Asc(FstChr(OStr)))
For J = 2 To Len(OStr)
    If IsPrintable <> IsAscPrintable(Asc(Mid(OStr, J, 1))) Then
        ShfTermzPrintable = Left(OStr, J - 1)
        OStr = Mid(OStr, J)
        Exit Function
    End If
Next
ShfTermzPrintable = OStr
OStr = ""
End Function

'Fun=================================================
Function LinRslt(EprLin, OvrFlwTerm$, S$) As LinRslt
With LinRslt
    .EprLin = EprLin
    .OvrFlwTerm = OvrFlwTerm
    .S = S
End With
End Function

Function EprzQuo$(BytAy() As Byte)
Dim O$(), I
For Each I In BytAy
    If I = vbDblQAsc Then PushI O, vb2DblQ Else PushI O, Chr(I)
Next
EprzQuo = QuoDbl(Jn(O))
End Function

Function EprzAndChr$(BytAy() As Byte)
Dim O$(), I
For Each I In BytAy
    PushI O, "Chr(" & I & ")"
Next
EprzAndChr = Jn(O, " & ")
End Function

Function Term(EprTerm$, S$) As Term
With Term
    .EprTerm = EprTerm
    .S = S
End With
End Function
Private Sub DqStrEpr__Tst()
Dim S$
GoSub YY1
GoSub YY2
GoSub T0
GoSub T1
Exit Sub
YY2:
    S = PjfStrP
    Brw DqStrEpr(S)
    Return
YY1:
    S = PjfStrP
    Brw DqStrEpr(S)
    Return
T0:
    S = "lksdjf lskdf dkf " & Chr(2) & Chr(11) & "ksldfj"
    Ept = Sy("")
    GoTo Tst
T1:
    GoTo Tst
Tst:
    Act = DqStrEpr(S)
    D Act
    Stop
    C
    Return
End Sub

Private Sub BrwRepeatedBytes__Tst()
BrwRepeatedBytes PjfStrP
End Sub

Function AscStr$(S)
Dim J&, O$()
For J = 1 To Len(S)
    PushI O, Asc(Mid(S, J, 1))
Next
AscStr = JnSpc(O)
End Function

Private Sub BrkAyzPrintable1__Tst()
Dim T, O$(), J&
'For Each T In BrkAyzPrintable(JnCrLf(SrcPth))
    J = J + 1
    Push O, FmtPrintableStr(T)
'Next
Brw AmAddIxPfx(O)
End Sub

Function FmtPrintableStr$(T)
Dim S$: S = PrintableSts(T)
Dim P$: P = S & " " & AliL(Len(T), 8) & " : "
Select Case S
Case "Prt": FmtPrintableStr = P & T
Case "Non": FmtPrintableStr = P & AscStr(Left(T, 10))
Case "Mix": FmtPrintableStr = P & AscStr(Left(T, 10))
Case Else
    Stop
End Select
End Function
Private Sub BrkAyzPrintable__Tst()
Brw BrkAyzPrintable(PjfStrP)
End Sub

Function BrkAyzRepeat(S) As String()
Dim L$: L = S
Dim T$, J&
While Len(L) > 0
    DoEvents
    T = ShfTermzRepeatedOrNot(L)
'    Debug.Print J, Len(L), Len(T), RepeatSts(T)
'    J = J + 1
    PushI BrkAyzRepeat, T
'    Stop
Wend
End Function
Function BrkAyzPrintable(S) As String()
Dim L$: L = S
#If True Then
    While Len(L) > 0
        Push BrkAyzPrintable, ShfTermzPrintable(L)
    Wend
#Else
    Dim T$, J&, I%
    While Len(L) > 0
        DoEvents
        T = ShfTermzPrintable(L)
        S = PrintableSts(T)
        Debug.Print J, Len(L), Len(T), S,
        If S = "NonPrintable" Then
            For I = 1 To Min(Len(T), 10)
                Debug.Print Asc(Mid(T, I, 1)); " ";
            Next
        End If
        Debug.Print
        
        J = J + 1
        PushI BrkAyzPrintable, T
    '    Stop
    Wend
#End If
End Function
Function PrintableSts$(T)
Dim IsPrintable As Boolean
IsPrintable = IsAscPrintable(Asc(FstChr(T)))
Dim J&
For J = 2 To Len(T)
    If IsPrintable <> IsAscPrintablezStrI(T, J) Then
        PrintableSts = "Mix"
        Stop
        Exit Function
    End If
Next
PrintableSts = IIf(IsPrintable, "Prt", "Non")
End Function

Function RepeatSts$(T)
'If Len(T) = 199 Then Stop
Select Case Len(T)
Case 0: RepeatSts = "ZeroByt": Exit Function
Case 1: RepeatSts = "OneByt":  Exit Function
Case Else
    Dim IsRepeat As Boolean, Las$, C$, IsSam As Boolean
    Las = SndChr(T)
    IsRepeat = FstChr(T) = Las
    Dim J&
    For J = 3 To Len(T)
        C = Mid(T, J, 1)
        IsSam = C = Las
        Select Case True
        Case IsRepeat And IsSam:
        Case IsRepeat: Stop: RepeatSts = "Mixed": Exit Function
        Case IsSam:    Stop: RepeatSts = "Mixed": Exit Function
        Case Else: Las = C
        End Select
    Next
End Select
RepeatSts = IIf(IsRepeat, "Repated", "Dif")
End Function
Function ShfTermzRepeatedOrNot$(OStr$)
If OStr = "" Then Exit Function
Dim J&, C$, Las$, IsSam As Boolean, IsRepeat As Boolean
Las = SndChr(OStr)
IsRepeat = FstChr(OStr) = Las
For J = 3 To Len(OStr)
    C = Mid(OStr, J, 1)
    IsSam = C = Las
    Select Case True
    Case IsSam And IsRepeat
    Case IsSam
        ShfTermzRepeatedOrNot = Left(OStr, J - 2)
        OStr = Mid(OStr, J - 1)
        Exit Function
    Case IsRepeat
        ShfTermzRepeatedOrNot = Left(OStr, J - 1)
        OStr = Mid(OStr, J)
        Exit Function
    Case Else
        Las = C
    End Select
Next
ShfTermzRepeatedOrNot = OStr
OStr = ""
End Function

Sub BrwRepeatedBytes(S)
Dim J&, B%, B1%, RepeatCnt&, L&
L = Len(S)
If L = 0 Then Exit Sub
B = Asc(FstChr(S)): RepeatCnt = 1
Erase XX
X FmtQQ("Len(?)", L)
For J = 2 To L
    B1 = Asc(Mid(S, J, 1))
    Select Case True
    Case B = B1:        RepeatCnt = RepeatCnt + 1
    Case Else
        If RepeatCnt > 1 Then
            X FmtQQ("Pos(?) Asc(?) RepeatCnt(?)", J, B, RepeatCnt)
            RepeatCnt = 1
        End If
        B = B1
    End Select
Next
Brw AmAddIxPfx(XX)
Erase XX
End Sub
