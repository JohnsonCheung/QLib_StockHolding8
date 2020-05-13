Attribute VB_Name = "MxVbStrFmt"
Option Compare Text
Option Explicit
Const CNs$ = "Str.Fmt"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbStrFmt."
Function MsgzDrs(MacroVbl$, MsgDta As Drs) As String()
Dim Fny$(): Fny = MsgDta.Fny
Dim Dr: For Each Dr In MsgDta.Dy
    PushI MsgzDrs, FmtMacroDi(MacroVbl, DiczFnyDr(Fny, Dr))
Next
End Function

Function FmtQQCrLf$(QQVbl$, ParamArray Ap())
FmtQQCrLf = FmtQQAv(QQVbl, Av) & vbCrLf
End Function

Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function FmtQQAv$(QQVbl$, Av())
Const CSub$ = CMod & "FmtQQAv"
Dim O$: O = Replace(QQVbl, "|", vbCrLf)
Dim P&: P = 1
Dim I: For Each I In Av
    P = InStr(P, O, "?")
    If P = 0 Then Exit For
    O = Left(O, P - 1) & Replace(O, "?", I, Start:=P, Count:=1)
    P = P + Len(I)
Next
FmtQQAv = O
End Function

Private Sub FmtQQAv__Tst()
Debug.Print FmtQQ("klsdf?sdf?dsklf", 2, 1)
End Sub

Function LblTabFmtAySepSS(Lbl$, Sy$()) As String()
PushI LblTabFmtAySepSS, Lbl
PushIAy LblTabFmtAySepSS, TabSy(Sy)
End Function

Function FmtV(V, Optional IsAddIx As Boolean) As String()
Select Case True
Case IsDic(V): FmtV = FmtDic(CvDic(V))
Case IsAet(V): FmtV = CvAet(V).Sy
Case IsLines(V): FmtV = AddIxPfxzLines(V)
Case IsPrim(V): FmtV = Sy(V)
Case IsSy(V)
    If IsAddIx Then
        FmtV = AmAddIxPfx(CvSy(V))
    Else
        FmtV = V
    End If
Case IsNothing(V): FmtV = Sy("#Nothing")
Case IsEmpty(V): FmtV = Sy("#Empty")
Case IsMissing(V): FmtV = Sy("#Missing")
Case IsObject(V): FmtV = Sy("#Obj(" & TypeName(V) & ")")
Case IsArray(V)
    Dim I, O$()
    If Si(V) = 0 Then Exit Function
    For Each I In V
        PushI O, StrfyVal(I)
    Next
    If IsAddIx Then
        FmtV = AmAddIxPfx(O)
    Else
        FmtV = O
    End If
Case Else
End Select
End Function
Function FmtStr$(DqStr$, ParamArray Ap())
Dim S$, J%
S = Replace(DqStr, "|", vbCrLf)
Dim I: For Each I In Ap
    S = Replace(S, "{" & J & "}", Nz(I, "Null"))
    J = J + 1
Next
FmtStr = S
End Function
