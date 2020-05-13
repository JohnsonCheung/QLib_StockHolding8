Attribute VB_Name = "MxDaoTbAttFfnOpEns"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxDaoTbAttFfnOpEns."
Public Const CNs$ = "Att"
Sub EnsAttFfn(D As Database, Attn$, Ffn$, Optional Attf$)
If W1IsAttNewer(D, Attn, Ffn) Then ImpAtt D, Attn, Ffn
If W1IsAttSam(D, Ffn) Then Exit Sub
ExpAttFfn D, Ffn
End Sub
Private Function W1IsAttSam(D As Database, Ffn$) As Boolean

End Function

Private Function W1IsAttNewer(D As Database, Attn$, Ffn$) As Boolean
If NoFfn(Ffn) Then Exit Function
W1IsAttNewer = FileDateTime(Ffn) > AttTim(D, Attn, Fn(Ffn))
End Function

Sub ExpAttFfn(D As Database, Ffn$)
Dim P$: P = Pth(Ffn)
Dim F$: F = Fn(Ffn)
ChkAttExist D, P, F
Dim A As Attd: A = Attd(D, P)        'Use Pth as Attn
End Sub
Sub ChkAttExist(D As Database, Attn$, Fn$, Optional Fun$ = "ChkAttExist")
If NoAtt(D, Attn, Fn) Then Thw Fun, "Att not exist", "Db Attn Fn", D.Name, Attn, Fn
End Sub
