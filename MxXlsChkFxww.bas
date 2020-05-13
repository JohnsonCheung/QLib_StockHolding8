Attribute VB_Name = "MxXlsChkFxww"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxXlsChkFxww."
Sub ChkFxww(Fx$, WW$, Optional Kd$ = "Excel file")
ChkFfnExist Fx, Kd
Dim Wny$(): Wny = WnyzFx(Fx)
Dim O$(): O = MinusSy(SyzSS(WW), Wny)
If Si(O) = 0 Then Exit Sub
Dim J%
Dim Er$()
Dim M$: M = FmtQQ("Missing ? worksheet", Si(O))
PushI Er, M
PushI Er, UL(M)
PushI Er, "Excel File    : [" & Fx & "]"
PushI Er, "Has worksheets: [" & Wny(0) & "]"
For J = 1 To UB(Wny)
PushI Er, "                [" & Wny(J) & "]"
Next
PushI Er, "Missing       : [" & O(0) & "]"
For J = 1 To UB(O)
PushI Er, "                [" & O(J) & "]"
Next
BrwEr Er
End Sub
