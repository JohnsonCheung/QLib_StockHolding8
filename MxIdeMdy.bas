Attribute VB_Name = "MxIdeMdy"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Ln"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdy."
Const C_Pub$ = "Public"
Const C_Prv$ = "Private"
Const C_Frd$ = "Friend"

Function AddPrv$(Ln, IsPrv As Boolean)
If IsPrv Then AddPrv = "Private " & Ln: Exit Function
AddPrv = Ln
End Function

Function ShtMdyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Pub", "Prv", "Frd", "")
ShtMdyAy = X
End Function

Function MdyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Pub, C_Prv, C_Frd, "")
MdyAy = X
End Function

Function Mdy$(Ln): Mdy = PfxzAySpc(Ln, MdyAy): End Function '#Modifier# :Nm FstTerm of a line (if it is Public | Private | Friend) otherwise *Blank

Function HitShtMdy(ShtMdy$, ShtVbMdyAy$()) As Boolean
HitShtMdy = HitAy(IIf(ShtMdy = "", "Pub", ShtMdy), ShtVbMdyAy)
End Function

Function ShtMdy$(Mdy)
Dim O$
Select Case Mdy
Case "Public", "": O = ""
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case Else: O = "???"
End Select
ShtMdy = O
End Function

Function IsMdy(S) As Boolean
IsMdy = HasEle(MdyAy, S)
End Function

Function RmvMdy$(Ln)
RmvMdy = LTrim(RmvPfxSySpc(Ln, MdyAy))
End Function
