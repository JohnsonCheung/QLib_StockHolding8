Attribute VB_Name = "MxVbAySrt"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Srt"
Const CMod$ = CLib & "MxVbAySrt."
Enum eSrtOrd
    eOrdAsc
    eOrdDes
End Enum

Function IxyzSrt(Ay, Optional Des As Boolean) As Long() ' ret Ixy of @Ay which is sorted
If Si(Ay) = 0 Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = InsBef(O, J, W2Ix(O, Ay, Ay(J)))
Next
If Des Then O = RevAy(O)
IxyzSrt = O
End Function

Private Function W2Ix&(Ix&(), Ay, V) ' ret the ix
Dim I, O&: For Each I In Ix
    If V < Ay(I) Then GoTo X
    O = O + 1
Next
X:
W2Ix = O
End Function

Function SrtLines$(A$)
SrtLines = JnCrLf(SrtAy(SplitCrLf(A)))
End Function

Function IsSrtd(Ay) As Boolean
Dim J&: For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Function
Next
IsSrtd = True
End Function

Private Sub SrtAyByAy__Tst()
Dim Ay, ByAy
Ay = Array(1, 2, 3, 4)
ByAy = Array(3, 4)
Ept = Array(3, 4, 1, 2)
GoSub Tst
Exit Sub
Tst:
    Act = SrtAyByAy(Ay, ByAy)
    C
    Return
End Sub

Function SrtAyByAy(Ay, ByAy)
Dim O: O = NwAy(Ay)
Dim I
For Each I In ByAy
    If HasEle(Ay, I) Then PushI O, I
Next
PushIAy O, MinusAy(Ay, O)
SrtAyByAy = O
End Function

'--
Private Sub SrtAy__Tst()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                   Act = SrtAy(A):        ChkEqAy Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = SrtAy(A):       ChkEqAy Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:HasPfx:Function"
Push A, "Private:MdMthDRsFunBdyLy:Function"
Push A, "Private:SrcMthLx_ToLx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:HasPfx:Function"
Push Exp, "Private:MdMthDRsFunBdyLy:Function"
Push Exp, "Private:SrcMthLx_ToLx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = QSrt(A)
ChkEqAy Exp, Act
End Sub

Function SrtAy(Ay, Optional By As eOrd)
If Si(Ay) = 0 Then SrtAy = Ay: Exit Function
Dim Ix&
Dim O: O = Ay: Erase O
PushI O, Ay(0)
Dim J&: For J = 1 To UB(Ay)
    Ix = W1InsBefIx(O, Ay(J))
    O = InsBef(O, Ay(J), Ix)
Next
If By = eByDes Then O = RevAy(O)
SrtAy = O
End Function

Private Function W1InsBefIx&(SrtdAy, V) ' ret an ix of @SrtdAy, so that @V should insert be that ix
Dim O&: For O = 0 To UB(SrtdAy)
    If SrtdAy(O) >= V Then W1InsBefIx = O: Exit Function
Next
W1InsBefIx = O
End Function

Private Sub IxyzSrt__Tst()
Dim A: A = Array("A", "B", "C", "D", "E")
ChkEqAy Array(0, 1, 2, 3, 4), IxyzSrt(A)
ChkEqAy Array(4, 3, 2, 1, 0), IxyzSrt(A, True)
End Sub

Function SrtAyInEIxIxy&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then SrtAyInEIxIxy& = O: Exit Function
        O = O + 1
    Next
    SrtAyInEIxIxy& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then SrtAyInEIxIxy& = O: Exit Function
    O = O + 1
Next
SrtAyInEIxIxy& = O
End Function

Private Sub SrtAy4__Tst()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                    Act = SrtAy(A):        ChkEq Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = SrtAy(A):       ChkEq Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:HasPfx:Function"
Push A, "Private:MdMthDRsFunBdyLy:Function"
Push A, "Private:SrcMthLx_ToLx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:HasPfx:Function"
Push Exp, "Private:MdMthDRsFunBdyLy:Function"
Push Exp, "Private:SrcMthLx_ToLx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = SrtAy(A)
ChkEq Exp, Act
End Sub

Private Sub IxyzSrtAy5__Tst()
Dim A: A = Array("A", "B", "C", "D", "E")
ChkEq Array(0, 1, 2, 3, 4), IxyzSrt(A)
ChkEq Array(4, 3, 2, 1, 0), IxyzSrt(A, True)
End Sub
