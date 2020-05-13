Attribute VB_Name = "MxIdeSrcCnstEns"
Option Explicit
Option Compare Text
Const CNs$ = "Md3Cnst"
Const CLib$ = "QIde"
Const CMod$ = CLib & "MxIdeSrcCnstEns."

Sub RmvCnst(M As CodeModule, Cnstn$)
Dim L&: L = CnstLno(M, Cnstn)
If L > 0 Then DltCdl M, CnstLno(M, Cnstn)
End Sub

Sub EnsCnst(M As CodeModule, CnstLin$)
'Ret : true if the const line is update, false if there is already such @CnstLin
Dim Lno&: Lno = CnstLno(M, Cnstn(CnstLin))
Select Case Lno
Case Is > 0
    If M.Lines(Lno, 1) = CnstLin Then Exit Sub
    RplCdl M, Lno, CnstLin
Case Else
    InsCnst M, CnstLin
End Select
End Sub

Sub EnsCnstAft(M As CodeModule, CnstLin$, AftCnstn$, Optional IsPrvOnly As Boolean)
Const CSub$ = CMod & "EnsCnstLinAft"
Dim Lno&: Lno = CnstLno(M, Cnstn(CnstLin))
If IsPrvOnly Then
    If Lno > 0 Then
        If HasPfx(M.Lines(Lno, 1), "Public ") Then
            Exit Sub
        End If
    End If
End If
If Lno > 0 Then
    If M.Lines(Lno, 1) = CnstLin Then Exit Sub
    M.ReplaceLine Lno, CnstLin
    InfLn CSub, "CnstLin is replaced", "Mdn CnstLin", Mdn(M), CnstLin
    Exit Sub
End If
InsCnstAft M, CnstLin, AftCnstn
End Sub

Sub InsCnst(M As CodeModule, CnstLin$)
InsCdl M, AftOptqImplLno(M), CnstLin
End Sub

Sub InsCnstAft(M As CodeModule, CnstLin$, AftCnstn$)
Const CSub$ = CMod & "InsCnstLinAft"
Dim Lno&
    Lno = CnstLno(M, AftCnstn): If Lno <> 0 Then Lno = Lno + 1
    If Lno = 0 Then Lno = AftOptqImplLno(M)
M.InsertLines Lno, CnstLin
InfLn CSub, "CnstLin is inserted", "Lno Mdn CnstLin", Lno, Mdn(M), CnstLin
End Sub

Sub ClrCnst(M As CodeModule, Cnstn$)
Const CSub$ = CMod & "ClrCnstLin"
Dim Lno&: Lno = CnstLno(M, Cnstn)
If Lno > 0 Then
    M.ReplaceLine Lno, ""
    InfLn CSub, "Cnstn is cleared", "Mdn Cnstn", Mdn(M), Cnstn
End If
End Sub

Sub RmvCnstLin(M As CodeModule, Cnstn$, Optional IsPrvOnly As Boolean)
Const CSub$ = CMod & "RmvCnstLin"
Dim Lno&: Lno = CnstLno(M, Cnstn, IsPrvOnly)
If Lno > 0 Then
    M.DeleteLines Lno, 1
    InfLn CSub, "Cnstn is removed", "Mdn Cnstn", Mdn(M), Cnstn
End If
End Sub

Sub RmvCnstLinzP(P As VBProject, Cnstn$, Optional IsPrvOnly As Boolean)
Dim C As VBComponent: For Each C In P.VBComponents
    RmvCnstLin C.CodeModule, Cnstn, IsPrvOnly
Next
End Sub
