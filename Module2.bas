Attribute VB_Name = "Module2"
Option Compare Text
Option Explicit
Sub AliSnglm(): AliSnglmzM CMd: End Sub
Private Sub AliSnglmzM(M As CodeModule):   RplMd M, AliSnglmSrcl(Src(M)):   End Sub
Private Function AliSnglmSrcl$(Src$())
Dim O$(): O = Src
Dim B() As Bei: B = SnglmBeiy(Src)
Dim J%: For J = 0 To BeiUB(B)
    RplAy O, AliBlk(AwBeiAsSy(Src, B(J))), B(J).Bix
Next
AliSnglmSrcl = JnCrLf(O)
End Function
Private Function SnglmBeiy(Src$()) As Bei(): SnglmBeiy = BeiyzBooly(IsSnglmBooly(Src)): End Function
Private Function IsSnglmBooly(Src$()) As Boolean()
Dim U&: U = UB(Src): If U = -1 Then Exit Function
Dim O() As Boolean: ReDim O(U)
Dim J%: For J = 0 To U
    If IsSngMthln(Src(J)) Then O(J) = True
Next
IsSnglmBooly = O
End Function
Private Function AliBlk(SnglmBlk$()) As String(): AliBlk = FmtDy(SnglmDy(SnglmBlk)): End Function
Private Function SnglmDy(SnglmBlk$()) As Variant()
Dim Snglm: For Each Snglm In SnglmBlk
    PushI SnglmDy, SnglmDr(Snglm)
Next
End Function
Private Function SnglmDr(Snglm) As String()
Dim O$(): O = SplitColon(Snglm)
With Brk2(O(1), "=")
SnglmDr = Sy(O(0), AddSfxIfNB(.S1, " ="), .S2, O(2))
End With
End Function
