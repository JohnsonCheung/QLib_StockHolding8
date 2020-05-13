Attribute VB_Name = "MxIdeMthLnIsCxt"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthlnIsCxt."
Const CNs$ = "AliMth"
Function IsMthlnCxt(L) As Boolean
'Ret : True ! if @Ln should be included as Mth-Context with one of is true
'           ! #1 IsRmk and aft (rmv ' and trim) not pfx <If Stop Insp == -- .. Brw>
'           ! #2 FstChr = :
'           ! #3 SngDimColon (&IsSngDimColon)   ! a dim and only one var and aft is [:]
'           ! #4 Is Asg stmt Ln (&IsLnAsg) @@
Dim Ln$: Ln = Trim(L)
Select Case True
Case HasPfx(L, "'")             ' Is Rmk
    Ln = LTrim(RmvFstChr(Ln))
    Select Case True
    Case HasPfxss(L, "If Stop Insp == -- .. Brw")     ' Don't incl if one of %Pfxy
    Case Else: IsMthlnCxt = True   ' <== Inl
    End Select
Case IsLnDimSngVarColon(Ln), IsLnAsg(Ln), FstChr(L) = ":"
    IsMthlnCxt = True              ' <== Inl
End Select
End Function

Function IsLnDimSngVarColon(L) As Boolean
'Ret true if L is Single-Dim-Colon: one V aft Dim and Colon aft DclSfx & not [For]
Dim Ln$: Ln = L
If Not ShfDim(Ln) Then Exit Function
If ShfNm(Ln) = "" Then Exit Function
ShfBkt Ln
ShfDclSfx Ln
'If HasSubStr(L, "For Each Dr In Itr(Dy") Then Stop
If FstChr(Ln) <> ":" Then Exit Function
If T1(RmvFstChr(Ln)) = "For" Then Exit Function '[Dim Dr: For ....] is False
IsLnDimSngVarColon = True
End Function

Private Sub IsLnDimSngVarColon__Tst()
Dim L
'GoSub T0
'GoSub T1
GoSub T3
'GoSub Z
Exit Sub
T3:
    L = "Dim Dr:       For JIsDesAy() As Boolean: IsDesAy = XIsDesAy(Ay)"
    Ept = False
    GoTo Tst
T1:
    L = "Dim IsDesAy() As Boolean: IsDesAy = XIsDesAy(Ay)"
    Ept = True
    GoTo Tst
T0:
    L = "Dim Aet As Access.Application: Set Aet = DftAcs(Acs)"
    Ept = True
    GoTo Tst
Tst:
    Act = IsLnDimSngVarColon(L)
    If Act <> Ept Then Stop
    Return
Z:
    Dim Aet As New Dictionary
    For Each L In SrczP(CPj)
        L = Trim(L)
        If T1(L) = "Dim" Then
            Dim S$: S = IIf(IsLnDimSngVarColon(L), "1", "0")
            Aet.PushItm S & " " & L
        End If
    Next
    VcAet Aet.Srt
    Return
End Sub

Function IsLnAsg(L) As Boolean
'Note: [Dr(NCol) = DicId(K)] is determined as Asg-Ln
Dim A$: A = LTrim(L)
ShfPfxSpc A, "Set"
If ShfDotNm(A) = "" Then Exit Function
If FstChr(A) = "(" Then
    A = AftBkt(A)
End If
IsLnAsg = T1(A) = "="
End Function
