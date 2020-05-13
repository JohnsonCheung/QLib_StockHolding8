Attribute VB_Name = "MxTpSpecPrp"
Option Explicit
Option Compare Text

Function SpeciHdrLixy(I() As Speci, Specit) As Integer() ' Return the Lix of @I which match the @Specit
Dim J%: For J = 0 To SpeciUB(I)
    If I(J).Specit = Specit Then PushI SpeciHdrLixy, I(J).Ix
Next
End Function

Function LyAyzSpeciy(I() As Speci) As Variant()
Dim J%: For J = 0 To SpeciUB(I)
    PushI LyAyzSpeciy, LyzILny(I(J).ILny)
Next
End Function

Function SpeciyzT(S As Spec, Specit$, Optional FmIx = 0) As Speci()
Dim I() As Speci: I = S.Itms
Dim J&: For J = FmIx To SpeciUB(S.Itms)
    Dim M As Speci: M = I(J)
    If M.Specit = Specit Then
        PushSpeci SpeciyzT, M
    End If
Next
End Function

Function Specity(I() As Speci) As String() '#spec-item-type-array#
Dim J%: For J = 0 To SpeciUB(I)
    PushI Specity, I(J).Specit
Next
End Function

