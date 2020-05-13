Attribute VB_Name = "MxIdeMthPmArgSfxDi"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthPmArgSfxDi."

':Vsfx: :Cml #Var-Sfx# ! It is from Arg or DimItm or Fun.  It is a sht form in direct attach to a :Var or :Argn or :Mthn
':Vn:   :Nm   #Var-Nm#
Function ArgSfxS12y(Mthly$()) As S12()
Dim L: For Each L In Itr(Mthly)
    PushS12 ArgSfxS12y, S12oDimnqVsfx(DimItmAyzS(Mthly))
Next
End Function

Function S12oDimnqVsfx(DimItm) As S12
Dim L$: L = DimItm
S12oDimnqVsfx.S1 = ShfNm(L)
S12oDimnqVsfx.S2 = Vsfx(L)
End Function

Function Vsfx$(VdclSfx$)
Const CSub$ = CMod & "Vsfx"
Dim S$: S = LTrim(VdclSfx)
Select Case True
Case S = "": Exit Function
Case Fst2Chr(S) = "()"
    If Len(S) = 2 Then
        Vsfx = S
    Else
        S = LTrim(Mid(S, 3))
        If HasPfx(S, "As ") Then
            Vsfx = ":" & Trim(RmvPfx(S, "As ")) & "()"
        Else
            Thw CSub, "Invalid VdclSfx", "When aft :() , it should be :As", "VdclSfx", VdclSfx
        End If
    End If
Case HasPfx(S, "As ")
    Vsfx = ":" & RmvPfx(RmvPfx(S, "As "), "New ")
Case Else
    Vsfx = S
End Select
End Function

Function S12oDimnqVsfxP() As S12()
S12oDimnqVsfxP = S12yoDimnqVsfxzP(CPj)
End Function

Function S12yoDimnqVsfxzP(P As VBProject) As S12()
S12yoDimnqVsfxzP = S12yoDimnqVsfx(DimItmAyzS(SrczP(P)))
End Function

Function S12yoDimnqVsfx(DimItmAy$()) As S12()
Dim I: For Each I In Itr(DimItmAy)
    PushS12 S12yoDimnqVsfx, S12oDimnqVsfx(I)
Next
End Function
