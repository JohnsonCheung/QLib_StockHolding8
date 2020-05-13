Attribute VB_Name = "MxIdeMthDrsMthl"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeMthDrsMthl."

Public Const MthlFF$ = "Mdn L Mthl" ' #Mth--Lines#

Sub ChkMthlDrs(Fun$, D As Drs)
ChkDrsFF Fun, "Mthl", D, MthlFF
End Sub

Function MthlDrs(MthcDrs As Drs) As Drs
MthlDrs = SelDrs(MthcDrs, "Mdn L Mthl")
End Function

Private Sub MthlDrszP__Tst()
BrwDrsN MthlDrszP(CPj)
End Sub

Private Sub MthlDrsP__Tst()
MthlDrsP
End Sub

Function MthlDrsP() As Drs
MthlDrsP = MthlDrszP(CPj)
End Function

Function MthlDrszP(P As VBProject) As Drs
Dim ODy()
    Dim C As VBComponent: For Each C In P.VBComponents
        PushIAy ODy, MthlDy(Src(C.CodeModule), C.Name)
    Next
MthlDrszP = DrszFF(MthlFF, ODy)
End Function

Function MthlDrszMthn(M As CodeModule, Mthn) As Drs
MthlDrszMthn = DrszFF(MthlFF, MthlDyzMthn(Src(M), Mdn(M), Mthn))
End Function

Function MthlDrszM(M As CodeModule) As Drs
MthlDrszM = DrszFF(MthlFF, MthlDy(Src(M), Mdn(M)))
End Function

Function MthlDy(Src$(), Mdn$) As Variant()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsMthln(Src(Ix)) Then
        PushI MthlDy, Array(Mdn, Ix + 1, Mthl(Src, Ix))
    End If
Next
End Function

Function MthlDyzMthn(Src$(), Mdn$, Mthn) As Variant()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsMthln(Src(Ix)) Then
        If Mthn(Src(Ix)) = Mthn Then
            PushI MthlDyzMthn, Array(Mdn, Ix + 1, Mthl(Src, Ix))
        End If
    End If
Next
End Function
