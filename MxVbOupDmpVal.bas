Attribute VB_Name = "MxVbOupDmpVal"
Option Explicit
Option Compare Text
Const CNs$ = "Vb.Val"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbOupDmpVal."

Sub D(V): DmpAy FmtV(V): End Sub
Sub Dmp(A): D A: End Sub
Sub DmpTyn(V): Debug.Print TypeName(V): End Sub ' Dmp Tyn of @V
Sub DmpAy(Ay, Optional WithIx As Boolean)  'Dmp Ay with Ix
Dim J&: For J = 0 To UB(Ay)
    DoEvents
    If WithIx Then Debug.Print J; ": ";
    Debug.Print Ay(J)
Next
End Sub

Sub OupAy(Ay, OupTy As eOupTy)
Select Case OupTy
Case eOupTy.eDmpOup: DmpAy Ay
Case eOupTy.eBrwOup: BrwAy Ay, "OupAy_"
Case eOupTy.eVcOup:  VcAy Ay, "VcAy_"
Case Else: PmEr "OupAy", "OupTy", OupTy, "0 1 2"
End Select
End Sub

