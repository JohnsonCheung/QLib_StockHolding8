Attribute VB_Name = "MxVbStrNB"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbStrNB."
Const CNs$ = "Str"
':FF: :Tml #Fldn-Spc-Sep# ! a list of Fldn has no space and separated by space.
Function AddNBAp$(ParamArray StrAp())
'Ret : :S ! ret a str by adding each ele of @StrAp one by one, if all them is <>'' else ret blank @@
Dim Av(): Av = StrAp
AddNBAp = JnNB(Av)
End Function

Function JnNB$(Ay, Optional Sep$ = ""): JnNB = Join(AwNB(Ay), Sep): End Function
Sub ChkIsNB(S, Fun$)
If Not IsNB(S) Then Thw Fun, "Given-S is not NB", "Trim(S)", Trim(S)
End Sub
Function AddSfxIfNB$(S_IfNB, Sfx$)
If IsNB(S_IfNB) Then AddSfxIfNB = S_IfNB & Sfx
End Function
Function AddPfxSpcIfNB(S_IfNB): AddPfxSpcIfNB = AddPfxIfNB(S_IfNB, " "): End Function
Function IsNB(S) As Boolean: IsNB = Trim(S) <> "": End Function
Function AddSfxDotIf(S_IfNB): AddSfxDotIf = AddSfxIfNB(S_IfNB, "."): End Function
Function AddSfxSpcIf(S_IfNB): AddSfxSpcIf = AddSfxIfNB(S_IfNB, " "): End Function

Function AddPfxIfNB$(IfNB_S, Pfx$)
If IfNB_S = "" Then Exit Function
AddPfxIfNB = Pfx & IfNB_S
End Function

Function AddPfxVbarIfNB$(S_IfNB)
AddPfxVbarIfNB = AddPfxIfNB(S_IfNB, " | ")
End Function

