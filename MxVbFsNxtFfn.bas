Attribute VB_Name = "MxVbFsNxtFfn"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsNxtFfn."

Private Sub NxtFfn__Tst()
Dim Ffn$
'GoSub T0
GoSub T1
Exit Sub
T1: Ffn = "AA(000).xls"
    Ept = "AA(001).xls"
    GoTo Tst
T0:
    Ffn = "AA.xls"
    Ept = "AA(001).xls"
    GoTo Tst
Tst:
    Act = NxtFfn(Ffn)
    C
    Return
End Sub

Function NxtFfn$(Ffn)
Dim J&: J = NxtNozFfn(Ffn)
Dim F$: F = RmvNxtNo(Ffn)
NxtFfn = AddFnSfx(F, "(" & Pad0(J + 1, 3) & ")")
End Function

Function NxtChdPth$(Pth)
':NxtChdPth: :Pth ! It is a child pth of @Pth with Fdr being :NxtSeqFdr
NxtChdPth = EnsPthSfx(Pth) & NxtFdr(Pth) & "\"
End Function

Function NxtFdr$(Pth)
Dim A$: A = MaxSeqFdr(Pth)
If A = "" Then
    NxtFdr = "0000"
Else
    NxtFdr = Pad0(Val(A) + 1, 4)
End If
End Function

Function MaxSeqFdr$(Pth)
Dim A$(): A = SeqEntAy(Pth): If Si(A) = 0 Then Exit Function
MaxSeqFdr = MaxEle(SeqEntAy(Pth))
End Function

Function SeqEntAy(Pth) As String()
':SeqFdr: :Fdr ! #Seq-Fdr-Ay# Fdr of name running 000 to 999
Dim A$(): A = EntAy(Pth, "????")
Dim I: For Each I In Itr(A)
    If IsNumeric(I) Then PushI SeqEntAy, I
Next
End Function

Function NxtFfnzNotIn(Ffn, NotInFfnAy$())
Dim J%, O$
O = Ffn
While HasStrEle(NotInFfnAy, O)
    LoopTooMuch CSub, J
    O = NxtFfn(O)
Wend
NxtFfnzNotIn = O
End Function

Function NxtFfnzAva$(Ffn)
Const CSub$ = CMod & "NxtFfnzAva"
Dim J%, O$
O = Ffn
While HasFfn(O)
    If J = 999 Then Thw CSub, "Too much next file in the path of given-ffn", "Given-Ffn", Ffn
    J = J + 1
    O = NxtFfn(O)
Wend
NxtFfnzAva = O
End Function

Function NxtFfnAy(Ffn) As String() 'Return ffn and all it nxt ffn in the pth of given ffn
If HasFfn(Ffn) Then Push NxtFfnAy, Ffn  '<==
Dim A$()
    Dim Spec$
        Spec = AddFnSfx(Fn(Ffn), "(???)")
    A = FfnAy(Pth(Ffn), Spec)
Dim I, F$
For Each I In Itr(A)
    F = I
    If IsNxtFfn(Ffn) Then PushI NxtFfnAy, F   '<==
Next
End Function

Function IsNxtFfn(Ffn) As Boolean
Select Case True
Case NxtNozFfn(Ffn) > 0, Right(Fnn(Ffn), 5) = "(000)": IsNxtFfn = True
End Select
End Function


Function NxtFn$(Fn$, FnAy$(), Optional MaxN% = 999)
If Not HasEle(FnAy, Fn) Then NxtFn = Fn: Exit Function
NxtFn = MaxEle(AwLik(FnAy, Fn & "(???)"))
End Function
