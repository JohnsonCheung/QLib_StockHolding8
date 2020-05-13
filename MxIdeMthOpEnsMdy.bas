Attribute VB_Name = "MxIdeMthOpEnsMdy"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthOpEnsMdy."

Sub EnsPrvMth(Mdn$, Mthn$)
'Ret : Ens a @Mthn in @Mdn as Private @@
If Not HasMd(CPj, Mdn) Then Exit Sub
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&: L = Mthlno(M, Mthn)
End Sub

Function EnsPrv$(Mthln)
Const CSub$ = CMod & "EnsPrv"
If Not IsMthln(Mthln) Then Thw CSub, "Given Mthln is not Mthln", "Ln", Mthln
EnsPrv = "Private " & RmvMdy(Mthln)
End Function


Function EnsPub$(Mthln)
Const CSub$ = CMod & "EnsPub"
If Not IsMthln(Mthln) Then Thw CSub, "Given Mthln is not Mthln", "Mthln", Mthln
EnsPub = RmvMdy(Mthln)
End Function

Sub EnsPjPrvZ()
EnsPrvYYP CPj
End Sub

Sub EnsPrvYYP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsPrvZ C.CodeModule
Next
End Sub

Sub EnsPrvZ(M As CodeModule, Optional Upd)
Const CmPfx$ = "X_"
Dim A As Drs: ' A = DPubZMth(M) ' L Mthln
Dim B As Drs: ' B = X_EnsPrv(A)   ' L Mthln PrvZ
Dim C As Drs: C = SelDrsAs(B, "L PrvZ:NewL Mthln:OldL")

RplLNewO M, C
End Sub

Function LnoAyOfPubZ(M As CodeModule) As Long()
Dim L, J&
For Each L In Itr(Src(M))
    J = J + 1
    'If IsMthlnzPub(L) Then
        PushI LnoAyOfPubZ, J
    'End If
Next
End Function

Function LnoItrOfPubZ(M As CodeModule)
Asg LnoItrOfPubZ, _
    Itr(LnoAyOfPubZ(M))
End Function

Function EnsMdy$(OldMthln, ShtMdy$)
Dim L$: L = RmvMdy(OldMthln)
    Select Case ShtMdy
    Case "Pub", "": EnsMdy = L
    Case "Prv":     EnsMdy = "Private " & L
    Case "Frd":     EnsMdy = "Friend " & L
    Case Else
        Thw "EnsMdy", "Given parameter [ShtMdy] must be ['' Pub Prv Frd]", "ShtMdy", ShtMdy
    End Select
End Function
