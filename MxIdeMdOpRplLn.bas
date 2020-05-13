Attribute VB_Name = "MxIdeMdOpRplLn"
Option Explicit
Option Compare Text
Const CNs$ = "Md.Ln.Op"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdRplLn."
Type PjLNewO
    Mdn As String
    LNewO As Drs
End Type
Public Const LNewOFF$ = "L NewL OldL"
Sub PushPjLNewO(O() As PjLNewO, M As PjLNewO)
Dim N&: N = PjLNewOSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub BrwPjLNewOAy(A() As PjLNewO)
BrwAy FmtPjLNewOAy(A), "PjNewOAy_"
End Sub
Function PjLNewO(Mdn$, LNewO As Drs) As PjLNewO
With PjLNewO
    .Mdn = Mdn
    .LNewO = LNewO
End With
End Function

Sub RplLNewO(M As CodeModule, LNewO As Drs, Optional Fun$ = "RplLNewO")
ChkIsLNewO LNewO, Fun
Dim Dr: For Each Dr In Itr(LNewO.Dy)
    Dim L&: L = Dr(0)
    Dim OldL$: OldL = Dr(2)
    Dim NewL$: NewL = Dr(1)
    If M.Lines(L, 1) <> OldL Then Thw Fun, "Md-Ln <> OldL", "Mdn Lno Md-Ln OldL NewL", Mdn(M), L, M.Lines(L, 1), OldL, NewL
    M.ReplaceLine L, NewL
Next
End Sub

Sub RplMdLnes(M As CodeModule, Lno&, OldLines$, NewLines$)
DltLines M, Lno, OldLines
M.InsertLines Lno, NewLines
End Sub

Sub ChkIsLNewO(LNewO As Drs, Optional Fun$ = "ChkIsLNewO")
If JnSpc(LNewO.Fny) <> LNewOFF Then Thw Fun, "Givn @Drs does have [L NewL OldL]", "FF-Drs-LNewO", JnSpc(LNewO.Fny)
End Sub
Function PjLNewOUB&(A() As PjLNewO)
PjLNewOUB = PjLNewOSi(A) - 1
End Function
Function PjLNewOSi&(A() As PjLNewO)
On Error Resume Next
PjLNewOSi = UBound(A) + 1
End Function
Function FmtPjLNewOAy(A() As PjLNewO) As String()
Dim J&: For J = 0 To PjLNewOUB(A)
    With A(J)
    PushI FmtPjLNewOAy, .Mdn
    PushIAy FmtPjLNewOAy, AmAddPfxTab(FmtLNewO(.LNewO))
    End With
Next
End Function

Function FmtLNewO(LNewO As Drs) As String()
Dim Dr: For Each Dr In Itr(LNewO.Dy)
    PushI FmtLNewO, Dr(0)
    PushI FmtLNewO, vbTab & Dr(1)
    PushI FmtLNewO, vbTab & Dr(2)
Next
End Function
Sub RplPjLNewO(P As VBProject, A() As PjLNewO)
Dim J%: For J = 0 To PjLNewOUB(A)
    With A(J)
        RplLNewO P.VBComponents(.Mdn).CodeModule, .LNewO
    End With
Next
End Sub
