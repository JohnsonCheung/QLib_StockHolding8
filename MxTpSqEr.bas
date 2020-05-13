Attribute VB_Name = "MxTpSqEr"
Option Explicit
Option Compare Text
#If False Then
'--SqTpEu
Private Type XEu
    A As String
End Type
Type SqEu
    A As String
    End Type ' Deriving(Ctor)

Type SqTpEr: Sw As SqSwEu: PmEu As SqPmEu: SqEu As SqEu: End Type 'Deriving(Ctor)
Function SqTpEr(S As SqTpSrc) As String()
Dim E As XEu
With E
      

End With
End Function
Private Function UUFmtEu(SqTp$(), E As XEu) As String()


End Function

Private Function SwlByNm(A() As Swl, SwNm) As Swl
Dim J%: For J = 0 To SwlUB(A)
    If A(J).Nm = SwNm Then SwlByNm = A(J): Exit Function
Next
End Function
Private Function SwNy(A() As Swl) As String()
Dim J%: For J = 0 To SwlUB(A)
    PushI SwNy, A(J).Nm
Next
End Function

Private Function SwlEr1(L As Swl, SwNm As Dictionary, Pm As Dictionary) As String()
'Each Term in A.Termy must be found either in Sw or Pm
Dim O0$(), O1$(), O2$(), I
For Each I In Itr(L.Term)
    Select Case True
    Case HasPfx(I, "?"):  If Not SwNm.Exists(I) Then Push O0, I
    Case HasPfx(I, "@?"): If Not Pm.Exists(I) Then Push O1, I
    Case Else:                  Push O2, I
    End Select
Next
SwlEr1 = AddAyAp( _
    TermNotInSwErn(O0, L, SwNm), _
    TermNotInPmErn(O1, L), _
    TermPfxEr(O2, L))
End Function

Private Function SwlMsg$(A As Swl, Msg$)

End Function

Private Function SwlEr(A As Swl) As String()

End Function

Private Function MustBeIntoLin$(A() As LLn)

End Function

Private Function MustBeSelorSelDis$(A() As LLn)

End Function

Private Function LeftOvrAftEvlEr(A() As Swl) As String()
'If Si(A) = 0 Then Exit Function
Dim I
PushI LeftOvrAftEvlEr, "Following lines cannot be further evaluated:"
'For Each I In A
'    PushI MsgzLeftOvrAftEvl, vbTab & CvSwl(I).Ln
'Next
'PushIAy MsgzLeftOvrAftEvl, FmtDicTit(Sw, "Following is the [Sw] after evaluated:")
End Function

Private Function SwNmMisErn$(A As Swl)
If A.Nm = "" Then
    SwNmMisErn = SwlMsg(A, "No sw name")
End If
End Function

Private Function SwOpErn$(A As Swl)
SwOpErn = SwlMsg(A, "2nd Term [Op] is invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Function


Private Function SwNmPfxErn$(A As Swl)
SwNmPfxErn = SwlMsg(A, "First Char of Sw-Nm must be @")
End Function

Private Function AndOrTermCntErn$(A As Swl)
If IsAndOr(A.Op) Then
    If Si(A.Term) <= 0 Then
        AndOrTermCntErn = SwlMsg(A, "[AND | OR]-Swl should have at least 1 term")
    End If
End If
End Function

Private Function EqNeTermCntErn$(A As Swl)
If IsEqNe(A.Op) Then
    If Si(A.Term) <> 2 Then
        EqNeTermCntErn = SwlMsg(A, "[Eq | Ne]-Swl should have exactly 2 terms")
    End If
End If
End Function


Private Function TermPfxEr(Termy, A As Swl) As String()
Dim T: For Each T In Itr(Termy)
    If Not HasSubStr("@?", FstChr(T)) Then
        PushI TermPfxEr, SwlMsg(A, "Terms[" & JnSpc(Termy) & "] must begin with either [?] or [@?]")
    End If
Next
End Function

Private Function TermNotInPmErn$(Termy, A As Swl)
TermNotInPmErn = SwlMsg(A, "Terms[" & JnSpc(Termy) & "] begin with [@?] must be found in Pm")
End Function

Private Function TermNotInSwErn$(Termy, A As Swl, SwNm As Dictionary)
TermNotInSwErn = SwlMsg(A, "Terms[" & JnSpc(Termy) & "] begin with [?] must be found in Switch")
End Function

Private Function DupSwNmEr(A() As Swl) As String()
DupSwNmEr = DupEr_(SywDup(SwNy(A)), A)
End Function

Private Function DupEr_(Dup$(), A() As Swl) As String()
Dim D: For Each D In Itr(Dup)
    PushI DupEr_, SwlMsg(SwlByNm(A, D), "There is dup name")
Next
End Function
#End If
