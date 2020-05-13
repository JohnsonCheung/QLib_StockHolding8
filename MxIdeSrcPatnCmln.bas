Attribute VB_Name = "MxIdeSrcPatnCmln"
Option Explicit
Option Compare Text
#If Doc Then
'Cml:Rule
'. #1 Fmt:/'xxx:Cml #aaa# /
'. #2 Must in #If-Doc-Blk-in-Dcl
'. #3 Must be one line
'. #4 Must be Cmln Memn [Description]
'CmlQ:Cml #Camel-Quote#
'Cmln:Cml #Camel-Name#
#End If
Enum eCmlnOpt: eExlCmlQ: eInlCmlQ: End Enum
'**Cmln
Private Sub Cmlny__Tst(): BrwAy Cmlny(SrclP, eExlCmlQ): End Sub
Function HasCmln(S) As Boolean: HasCmln = HasRx(S, CmlnRx): End Function
Function Cmln$(S): Cmln = MchszR(S, CmlnRx): End Function ' #Camel-name# a name between 2 hashChr
Function CmlnRx() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("'[A-Za-z][\w\.-]*:Cml ", IsGlobal:=True)
Set CmlnRx = X
End Function
Function Cmlny(S$, Optional H As eCmlnOpt) As String()
Cmlny = MchsyzR(S, CmlnRx)
If H = eExlCmlQ Then Cmlny = W1RmvCmlQzAy(Cmlny)
End Function
Function W1RmvCmlQzAy(Cmlny$()) As String()
Dim Cml: For Each Cml In Itr(Cmlny)
    PushI W1RmvCmlQzAy, W1RmvCmlQ(Cml)
Next
End Function
Private Sub W1RmvCmlQ__Tst(): MsgBox W1RmvCmlQ("'CC:Cml "): End Sub
Function W1RmvCmlQ$(CmlPatn): W1RmvCmlQ = Mid(CmlPatn, 2, Len(CmlPatn) - 6): End Function

'**CmlS12
Private Sub CmlS12y__Tst(): BrwS12y CmlS12y(SrcP): End Sub
Function CmlS12(Ln) As S12
Dim N$: N = Cmln(Ln): If N = "" Then Exit Function
Dim M$: M = Memn(Ln): If M = "" Then Thw CSub, "@Ln has Cmln, but no memn", "@Ln Cmln", Ln, N
CmlS12 = S12(N, M)
End Function
Function CmlS12y(Src$()) As S12()
Dim L: For Each L In Itr(Src)
    PushS12Opt CmlS12y, CmlS12(L)
Next
End Function

'**ErCmlJrc
Function ErCmlJrc() As String()
End Function
Private Function ErCmlJrczM(S As MdSrc) As String()
Dim Src$(): Src = S.Src
Dim Mdn$: Mdn = S.Mdn
Dim J&: For J = 0 To UB(Src)
    Dim L$: L = Src(J)
    If IsErCmlln(L) Then
        Dim P As C12: P = C12zRx(L, CmlnRx)
        PushI ErCmlJrczM, JrclnzP(Mdn, J + 1, P, L)
    End If
Next
End Function
Function IsErCmlln(Ln$) As Boolean
Select Case True
Case Not HasCmln(Ln)
Case Not HasMemn(Ln): IsErCmlln = True
End Select
End Function
