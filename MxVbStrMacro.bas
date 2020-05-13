Attribute VB_Name = "MxVbStrMacro"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrMacro."
Type NNAv
    NN As String
    Av() As Variant
End Type

Function MacroNy(Macro, Optional OpnBkt$ = vbBigOpn, Optional InlBkt As Boolean) As String()
'Macro is a str with ..[xx].., it is to return all xx or [xx]
Dim Q1$:   Q1 = OpnBkt
Dim Q2$:   Q2 = ClsBkt(OpnBkt)
Dim Sy$(): Sy = Split(Macro, Q1)
Dim O$():   O = AwDis(AeBlnk(BefzSy(Sy, Q2)))
If InlBkt Then O = AmAddPfxSfx(O, Q1, Q2)
MacroNy = O
End Function

Function RplMacro(MacroVbl, NN$, ParamArray ValAp())
Dim O$
    O = RplVbl(MacroVbl)
    Dim J%
    Dim Nm: For Each Nm In Itr(SyzSS(NN))
        Dim V: V = ValAp(J)
        O = Replace(O, "{" & Nm & "}", V)
        J = J + 1
    Next
RplMacro = O
End Function

Function FmtMacroDi$(MacroVbl, D As Dictionary)
Dim O$: O = RplVBar(MacroVbl)
Dim K: For Each K In D.Keys
    O = Replace(O, QuoBig(K), D(K))
Next
FmtMacroDi = O
End Function

Function FmtMacro$(MacroVbl, ParamArray Nap())
Dim Nav():  Nav = Nap
FmtMacro = FmtMacroNav(MacroVbl, Nav)
End Function

Function FmtMacroNav$(MacroVbl, Nav()): FmtMacroNav = FmtMacroDi(MacroVbl, DizNav(Nav)): End Function

Function FmtMacroRs$(Macro, Rs As DAO.Recordset)
FmtMacroRs = FmtMacroDi(Macro, DizRs(Rs))
End Function

Function DizRs(A As DAO.Recordset) As Dictionary
Set DizRs = New Dictionary
Dim F As DAO.Field
For Each F In A.Fields
    DizRs.Add F.Name, F.Value
Next
End Function

Function DizNav(Nav()) As Dictionary
Set DizNav = New Dictionary
If Si(Nav) > 0 Then
    Dim Ny$(): Ny = SyzSS(Nav(0))
    Dim J%: For J = 1 To Si(Ny)
        DizNav.Add Ny(J - 1), Nav(J)
    Next
End If
End Function

Function NNAv(NN$, Av()) As NNAv
Const CSub$ = CMod & "NNAv"
Dim N$(): N = Ny(NN)
ChkNy N, CSub
If Si(N) <> Si(Av) Then Thw CSub, "NN-Si <> Av-Si", "NN-Si Av-Si NN", Si(N), Si(Av), NN
NNAv.NN = NN
NNAv.Av = Av
End Function

Function NNAvzDic(A As Dictionary) As NNAv
NNAvzDic.NN = JnSpc(NyzItr(A.Keys))
NNAvzDic.Av = AvzItr(A.Items)
End Function
