Attribute VB_Name = "MxVbDtaLikss"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbDtaLikss."
Function IsLikzSS(A, Kss) As Boolean
IsLikzSS = IsLikzAy(A, SyzSS(Kss))
End Function

Function IsLikzAy(A, LikAy$()) As Boolean
Dim I
For Each I In Itr(LikAy)
    If A Like I Then IsLikzAy = True: Exit Function
Next
End Function

Function IsLikzSSAy(A, KssAy) As Boolean
Dim Kss
For Each Kss In KssAy
    If IsLikzSS(A, Kss) Then IsLikzSSAy = True: Exit Function
Next
End Function

Private Sub T1zTkssLy__Tst()
Dim A$(), Nm$
GoSub T1
GoSub T2
Exit Sub
T1:
    A = SplitVBar("a bb* *dd | c x y")
    Nm = "x"
    Ept = "c"
    GoTo Tst
T2:
    A = SplitVBar("a bb* *dd | c x y")
    Nm = "bb1"
    Ept = "a"
    GoTo Tst
Tst:
    Act = T1zTkssLy(A, Nm)
    C
    Return
End Sub

Function T1zTkssLy$(TkssLy$(), Nm)
':Tkss: :SS #T1-Likss# ! It is SS with T1 and Likss
':Kss:  :SS #Likss#    ! It is SS with each term is LikStr
Dim L: For Each L In Itr(TkssLy)
    Dim T1$: T1 = ShfTerm(L)
    If IsLikzSS(Nm, L) Then
        T1zTkssLy = T1
        Exit Function
    End If
Next
End Function
