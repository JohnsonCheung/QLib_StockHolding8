Attribute VB_Name = "MxIdeMthLnSngml"
Option Compare Database
Sub AliSngml(): AliSngmlzM CMd: End Sub
Sub AliSngmlzM(M As CodeModule): MdyMd M, SngmlAliSrc(Src(M)): End Sub
Function SngmlAliSrc(Src$()) As String()
Dim O$(): O = Src
Dim B() As Bei: B = BeiyzBooly(W1IsSnglmy(Src))
Dim J%: For J = 0 To BeiUB(B)
    RplAy O, W1Ali(AwBei(O, B(J))), B(J).Bix
Next
SngmlAliSrc = O
End Function
Private Function W1IsSnglmy(Src$()) As Boolean()
Dim L: For Each L In Itr(Contly(Src))
    PushI W1IsSnglmy, IsSngMthln(L)
Next
End Function
Private Function W1Ali(Snglmy$()) As String(): W1Ali = FmtStrColy(W1StrColy(Snglmy)): End Function
Private Function W1StrColy(Snglmy$()) As StrColy: W1StrColy = StrColyzDy(W1Dy(Snglmy)):    End Function
Private Function W1Dy(Snglmy$()) As Variant()
Dim L: For Each L In Itr(Snglmy)
    PushI W1Dy, W1Dr(L)
Next
End Function
Private Function W1Dr(Snglm) As String()
Dim L$: L = Snglm
Dim K$: K = MthKdzL(Ln)
Dim IsSub As Boolean: IsSub = K = "Sub"
PushI W1Dr, ShfBef(L, ":", eInlSep)
PushI W1Dr, W1ShfLHS(L, IsSub)
PushI W1Dr, ShfBef(L, "End " & K)
PushI W1Dr, L
End Function
Private Function W1Shf1$(OLn$): W1Shf1 = ShfBef(OLn, Bef, eInlSep): End Function
Private Function W1Shf2$(OLn$, IsSub As Boolean)
If Not IsSub Then W1ShfLHS = ShfBef(OLn, "=")
End Function
