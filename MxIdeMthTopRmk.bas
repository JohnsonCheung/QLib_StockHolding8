Attribute VB_Name = "MxIdeMthTopRmk"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthMrmk."
Sub MthBieyzSNT__Tst()
Dim Src$(), Mthn
Dim Ept() As Bei, Act() As Bei

Src = SrczMdn("MxMrmk")
PushBei Ept, Bei(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MthBieyzN(Src, Mthn)
    If Not IsEqBeiy(Act, Ept) Then Stop
    Return
End Sub

Function Mrmkl$(Src$(), Mthix): Mrmkl = JnCrLf(Mrmk(Src, Mthix)): End Function
Function Mrmk(Src$(), Mthix) As String()
Dim Fm&: Fm = Mrmkix(Src, Mthix): If Fm = -1 Then Exit Function
Mrmk = AeBlnk(AwBE(Src, Fm, Mthix - 1))
End Function
Function Mrmkix&(Src$(), Mthix)
If Mthix <= 0 Then Mrmkix = -1: Exit Function
Dim J&, L$, I&
Mrmkix = Mthix
For J = Mthix - 1 To 0 Step -1
    If Not IsRmkOrBlnk(Src(J)) Then
        For I = J To Mthix
            If Not IsBlnk(Src(I)) Then Mrmkix = I: Exit Function
        Next
        Imposs CSub
    End If
    L = LTrim(Src(J))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": Mrmkix = J
    Case Else: Exit Function
    End Select
Next
End Function

Function MrmkLno(Md As CodeModule, Mthlno)
Dim J&, L$
MrmkLno = Mthlno
If Mthlno = 0 Then Exit Function
For J = Mthlno - 1 To 1 Step -1
    L = LTrim(Md.Lines(J, 1))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": MrmkLno = J
    Case Else: Exit Function
    End Select
Next
End Function
Private Sub EnsMrmk__Tst()
'GoSub Z1
Dim M As CodeModule
GoSub Z1
Exit Sub
Z1:
    'GoSub Crt: Exit Sub
    Set M = Md("TmpMod123")
    EnsMrmk M, "AAXX", "'skldfjsdlkfj lksdj flksdj fkj @@"
    Return
Z2:
    Set M = Md("TmpMod20190605_231101")
    EnsMrmk M, "AAXX", RplVBar("'sldkfjsd|'slkdfj|slkdfj|'sldkfjsdf|'sdf")
    Return
Z3:
    Set M = Md("TmpMod20190605_231101")
    EnsMrmk M, "AAXX", RplVBar("'a|'bb|'cfsdfdsc")
    Return
Crt:
    EnsMod CPj, "TmpMod123"
    Set M = Md("TmpMod123")
    ClrMd M
    M.AddFromString "Sub AAXX()" & vbCrLf & "End Sub"
    Return
End Sub
Sub EnsMrmk(M As CodeModule, Mthn, Mrmkl$)

End Sub
