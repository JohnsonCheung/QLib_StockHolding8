Attribute VB_Name = "MxDtaDaFmt"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaFmt."
Enum eTblFmt: eNoSep: eColSep: eRowSep: eBothSep: End Enum
Public Const eTblFmtSS$ = "NoSep ColSep RowSep BothSep"
Type DrsFmto ' Drs format option
    MaxWdt As Integer 'Max-Column-Wdt
    Brkcc As String   'Breaking-column-cc
    BegIx As Integer
    ShwZer As Boolean
    Fmt As eTblFmt
    IsSum As Boolean
End Type

Function FmtDs(D As Ds, Optional DrsFmts$) As String()
FmtDs = FmtDsO(D, DrsFmtozS(DrsFmts))
End Function

Function FmtDsO(D As Ds, Opt As DrsFmto) As String()
PushI FmtDsO, "*Ds " & D.DsNm & " " & String(10, "=")
Dim Dic As Dictionary
    Set Dic = DiczVbl(Opt.Brkcc)
Dim J%: For J = 0 To DtUB(D.DtAy)
    PushAy FmtDsO, FmtDtO(D.DtAy(J), Opt)
Next
End Function
Private Sub FmtDrsN__Tst()
Dim A As Drs, DrsFmts$
GoSub Z
Exit Sub
T1:
    A = SampDrs
    GoSub Tst
Tst:
    Act = FmtDrsN(A, DrsFmts)
    Brw Act: Stop
    C
    Return
Z:
    DmpAy FmtDrsN(SampDrs1, DrsFmts)
    Return
End Sub

Function FmtDrsN(D As Drs, Optional DrsFmts$) As String() ' Format normally without option.  Normally = without reduce)
FmtDrsN = FmtDrsNO(D, DrsFmtozS(DrsFmts))
End Function

Function FmtDrsNO(D As Drs, Opt As DrsFmto) As String() ' Format normally with option.  Normally = without reduce)
'@BrkColnn : if changed, insert a break line if BrkColNm is given
If NoReczDrs(D) Then FmtDrsNO = ZZNoRecMsg(D): Exit Function
Dim WiIxCol As Drs: WiIxCol = InsIxColzDrs(D, Opt.BegIx)
Dim WiSepDr As Drs: WiSepDr = W1AddBrkDr(WiIxCol, Opt.Brkcc)
Dim Dy():                Dy = AddAyEle(WiSepDr.Dy, WiSepDr.Fny)

Dim Bdy$():       Bdy = FmtDy(Dy, Opt.MaxWdt, Opt.Fmt)

Dim Sep$:         Sep = Pop(Bdy)                                ' Sep-Ln
Dim Hdr$:         Hdr = Pop(Bdy)                                ' Hdr-Ln
              FmtDrsNO = Sy(Sep, Hdr, Bdy, Sep)
End Function

Private Function W1AddBrkDr(D As Drs, Brkcc$) As Drs
'Adding BrkDr to @A:SepDr: :Dr ! #Separated-Dr# for a Drs.  It is an Empty :Dr inserted between rows for a Drs for separation
Select Case True
Case Brkcc = "", NoReczDrs(D): W1AddBrkDr = D: Exit Function
End Select

Dim ShdBrk() As Boolean: ShdBrk = W1ShdBrkAy(D, Brkcc)

If Si(ShdBrk) = 0 Then W1AddBrkDr = D: Exit Function
Dim Dy(): Dy = D.Dy
Dim ODy()
    Dim J&: For J = 0 To UB(Dy)
        If ShdBrk(J) Then
            PushI ODy, Array() '<==
        End If
        PushI ODy, Dy(J)
    Next
W1AddBrkDr = Drs(D.Fny, ODy)
End Function

Private Function W1ShdBrkAy(D As Drs, Brkcc$) As Boolean() ' Should-break-array, which is same size as Drs-@D means to insert a blank record before that record
Dim Dy(): Dy = D.Dy
Dim Ixy&(): Ixy = IxyzCC(D, Brkcc)
Dim LasK: LasK = AwIxy(Dy(0), Ixy)
Dim CurK
Dim Dr: For Each Dr In Itr(Dy)
    CurK = AwIxy(Dr, Ixy)
           PushI W1ShdBrkAy, Not IsEqAy(CurK, LasK)
    LasK = CurK
Next
End Function

'---===================================
Function FmtDtO(A As Dt, Opt As DrsFmto) As String()
PushI FmtDtO, "*Tbl " & A.DtNm
PushIAy FmtDtO, FmtDrsRO(DrsFmDt(A), Opt)
End Function

Function FmtDt(D As Dt, Optional DrsFmts$) As String()
FmtDtO D, DrsFmtozS(DrsFmts)
End Function

Private Sub FmtDt__Tst()
Dim A As Dt, Opt As DrsFmto
'--
A = SampDt1
'Ept = Z_TimStrpt1
GoSub Tst
'--
Exit Sub
Tst:
    Act = FmtDtO(A, Opt)
    C
    Return
End Sub


Function FmtDrsV(D As Drs, Optional Nm$) As String() ' format Drs @D as vertical format
If NoReczDrs(D) Then PushI FmtDrsV, ZZNoRecMsg(D, Nm): Exit Function
Dim Fny$(): Fny = AmAli(AmAddIxPfx(D.Fny))
Dim N&: N = NReczDrs(D)
Dim J&: For J = 0 To UB(D.Dy)
    Dim Dr(): Dr = D.Dy(J)
    PushIAy FmtDrsV, FmtDrV(Dr, Fny, J, N)
Next
End Function

Function FmtDrV(Dr(), Fny$(), Optional Ix& = -1, Optional NRec&) As String()
If Ix >= 0 Then PushI FmtDrV, "Record Ix: " & Ix & " of " & NRec
PushIAy FmtDrV, FmtNyAv(Fny, Dr)
PushI FmtDrV, ""
End Function

'---=========================
Function DftDrsFmto() As DrsFmto
DftDrsFmto = DrsFmto
End Function

Function DrsFmtozS(DrsFmts$) As DrsFmto
With DrsFmtozS
Dim N$(): N = Termy(DrsFmts)
If ShfEle(N, "ShwZer") Then .ShwZer = True
Dim W%: W = ShfIntBet(N, 3, 1200): If W <> 0 Then .MaxWdt = W
.IsSum = ShfEle(N, "IsSum")
If ShfEle(N, "TblFmt") Then
    .Fmt = eNoSep
ElseIf ShfEle(N, "SSFmt") Then
    .Fmt = eColSep
End If
End With
If Si(N) <> 0 Then W1Warn N, DrsFmts
End Function

Private Sub W1Warn(N$(), DrsFmts$) ' N should not have anything left


End Sub

Function DrsFmto(Optional MaxWdt% = 100, Optional Brkcc$, Optional ShwZer As Boolean, Optional IsSum As Boolean, Optional BegIx% = 1, Optional Fmt As eTblFmt) As DrsFmto
With DrsFmto
    .MaxWdt = MaxWdt
    .Brkcc = Brkcc
    .ShwZer = ShwZer
    .BegIx = BegIx
    .Fmt = Fmt
    .IsSum = IsSum
End With
End Function

Private Function ZZNoRecMsg(D As Drs, Optional Nm$ = "Drs1") As String()
Dim FF$
FF = JnSpc(D.Fny)
If FF = "" Then FF = " (No Fny)"
FF = FmtQQ("Drs(?) (NoRec) ?", Nm, FF)
ZZNoRecMsg = Sy(FF)
End Function

'Move to other modules ===========================================
Function StrfyVal$(V, Optional W% = 100)
Dim O$
Select Case True
Case IsSy(V):      O = StrfySy(CvSy(V), W)
Case IsStr(V):     O = V
Case IsBool(V):    O = IIf(V, "True", "")
Case IsPrim(V):    O = V
Case IsEmp(V):     O = ""
Case IsNull(V):    O = ""
Case IsPrimy(V):  O = StrfyPrimAy(V)
Case IsDic(V):     O = "#Dic:Cnt(" & CvDic(V).Count & ")"
Case IsObject(V):  O = "#O:" & TypeName(V)
Case IsErObj(V):   O = "#Er#"
Case Else:         O = V
End Select
StrfyVal = O
End Function
Function StrfySy$(Sy$(), Optional W% = 100)
If Si(Sy) = 0 Then Exit Function
StrfySy = Left(Sy(0), W)
End Function
Function StrfyPrimAy$(V, Optional W% = 100)
StrfyPrimAy = "*[" & Si(V) & "]"
End Function
Function LinesifySy(Sy$(), W%)
Dim Ln: For Each Ln In Itr(Sy)
    PushIAy LinesifySy, WrpLn(Ln, W)
Next
End Function

Function SslSyzDy(Dy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    Push SslSyzDy, SslzDr(Dr) ' Fmtss(X)
Next
End Function


Function IsEqAyzIxy(A, B, Ixy&()) As Boolean
Dim J%
For J = 0 To UB(Ixy)
    If A(Ixy(J)) <> B(Ixy(J)) Then Exit Function
Next
IsEqAyzIxy = True
End Function

Function JnDySpc(Dy()) As String()
JnDySpc = JnDy(Dy, " ")
End Function

Function JnDyDot(Dy()) As String()
JnDyDot = JnDy(Dy, ".")
End Function

Function JnDy(Dy(), Optional Sep$ = " ") As String()
'Ret: :Ly by joining each :Dr in @Dy by @Sep
Dim Dr: For Each Dr In Itr(Dy)
    PushI JnDy, Jn(Dr, Sep)
Next
End Function
