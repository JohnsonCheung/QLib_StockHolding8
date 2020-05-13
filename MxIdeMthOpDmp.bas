Attribute VB_Name = "MxIdeMthOpDmp"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxIdeMthOpDmp."
Enum eWhMdMth: eAllMdMth: eOnlyPub: ePubAndTst: End Enum
Const eWhMdMthSS$ = "AllMdMth OnlyPub PubAndTst"
Private Type XWh: InlPrv As Boolean: InlTst As Boolean: InlFrd As Boolean: End Type

Sub DmpMdPfx(MdnPfx$, Optional Wh As eWhMdMth)
Dmp FmtMdPfx(MdnPfx, Wh)
End Sub

Function FmtMdPfx(MdnPfx$, Optional Wh As eWhMdMth) As String()
FmtMdPfx = XFmt(MdnyzPfx(MdnPfx), Wh)
End Function

Sub DmpMdSubStr(MdnSubStr$, Optional Wh As eWhMdMth)
Dmp FmtMdSubStr(MdnSubStr, Wh)
End Sub

Function FmtMdSubStr(MdnSubStr$, Optional Wh As eWhMdMth) As String()
FmtMdSubStr = XFmt(MdnyBySubStr(MdnSubStr), Wh)
End Function

Private Function XFmt(Mdny$(), Wh As eWhMdMth) As String()
Dim O() As S12
Dim W As XWh: W = W1CvWh(Wh)
Dim Mdn: For Each Mdn In Itr(Mdny)
    PushS12Ay O, W1OneMd(Md(Mdn), W)
Next
XFmt = FmtS12y(O, "Mdn Mthn")
End Function

Private Function W1CvWh(Wh As eWhMdMth) As XWh
With W1CvWh
Select Case True
Case Wh = eAllMdMth: .InlPrv = True: .InlFrd = True: .InlTst = True
Case Wh = eOnlyPub
Case Wh = ePubAndTst: .InlTst = True
Case Else: EnmEr CSub, "eWhMdMth", eWhMdMthSS, Wh
End Select
End With
End Function

Private Function W1OneMd(M As CodeModule, W As XWh) As S12() ' Return the S12() format from one @M by @W
Dim O() As S12
Dim N() As Mthn3: N = Mthn3yzM(M)
Dim PubTst As S12: PubTst = W1ShfPubTst(N, "  PubTst")
Dim PrvTst As S12: PrvTst = W1ShfPrvTst(N, "  PrvTst")
Dim Pub As S12: Pub = W1ShfPub(N, Mdn(M))
Dim Frd As S12: Frd = W1ShfPrv(N, "  Prv")
Dim Prv As S12: Prv = W1ShfFrd(N, "  Frd")
If S12Si(O) > 0 Then Imposs CSub, "After shifting Tst Pub Frd Prv, there is still something left in Mthn3y of a module."
PushS12 O, Pub
With W
    W1Push O, .InlTst, PubTst
    W1Push O, .InlTst, PrvTst
    W1Push O, .InlFrd, Frd
    W1Push O, .InlPrv, Prv
End With
W1OneMd = O
End Function

Private Sub W1Push(O() As S12, B As Boolean, M As S12)
If Not B Then Exit Sub
If M.S2 = "" Then Exit Sub
PushS12 O, M
End Sub
Private Function W1IsTst(Mthn) As Boolean
Select Case True
Case HasPfx(Mthn, "T_"), HasSfx(Mthn, "__Tst"): W1IsTst = True
End Select
End Function
Private Function W1IsPrvTst(Mthn, ShtMdy$) As Boolean
If ShtMdy <> "Prv" Then Exit Function
W1IsPrvTst = W1IsTst(Mthn)
End Function

Private Function W1IsPubTst(Mthn, ShtMdy$) As Boolean
If ShtMdy <> "" Then Exit Function
W1IsPubTst = W1IsTst(Mthn)
End Function

Private Function W1ShfPub(O() As Mthn3, S1$) As S12
Dim N$(), OO() As Mthn3
Dim M As Mthn3
Dim J&: For J = 0 To Mthn3UB(O)
    M = O(J)
    If M.ShtMdy = "" Then
        PushI N, M.Nm
    Else
        PushMthn3 OO, M
    End If
Next
O = OO
W1ShfPub = XS12(S1, N)
End Function

Private Function W1ShfPubTst(O() As Mthn3, S1$) As S12
Dim OO() As Mthn3
Dim M As Mthn3, N$()
Dim J&: For J = 0 To Mthn3UB(O)
    M = O(J)
    If W1IsPubTst(M.Nm, M.ShtMdy) Then
        PushI N, M.Nm
    Else
        PushMthn3 OO, M
    End If
Next
W1ShfPubTst = XS12(S1, N)
O = OO
End Function

Private Function W1ShfPrvTst(O() As Mthn3, S1$) As S12
Dim OO() As Mthn3
Dim M As Mthn3, N$()
Dim J&: For J = 0 To Mthn3UB(O)
    M = O(J)
    If W1IsPrvTst(M.Nm, M.ShtMdy) Then
        PushI N, M.Nm
    Else
        PushMthn3 OO, M
    End If
Next
W1ShfPrvTst = XS12(S1, N)
O = OO
End Function

Private Function W1ShfPrv(O() As Mthn3, S1$) As S12
Dim OO() As Mthn3
Dim M As Mthn3, N$()
Dim J&: For J = 0 To Mthn3UB(O)
    M = O(J)
    If M.ShtMdy = "Prv" Then
        PushI N, M.Nm
    Else
        PushMthn3 OO, M
    End If
Next
W1ShfPrv = XS12(S1, N)
O = OO
End Function

Private Function W1ShfFrd(O() As Mthn3, S1$) As S12
Dim OO() As Mthn3
Dim M As Mthn3, N$()
Dim J&: For J = 0 To Mthn3UB(O)
    M = O(J)
    If M.ShtMdy = "Frd" Then
        PushI N, M.Nm
    Else
        PushMthn3 OO, M
    End If
Next
W1ShfFrd = XS12(S1, N)
O = OO
End Function
'== X
Function XS12(S1$, N$()) As S12
If HasEle(N, "FstAttFn") Then Stop
XS12 = S12Trim(S1, JnSpc(QSrt(N)), NoTrim:=True)
End Function
