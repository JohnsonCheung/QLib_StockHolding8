Attribute VB_Name = "MxIdeMthMsigPm"
Option Explicit
Option Compare Text
Const CNs$ = "Mth"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthMsigPm."

Function MthPm$(Mthln)
':MthPm: :Str #Mth-Parameters#
MthPm = BetBkt(Mthln)
End Function

Function MthPmAy(Mthlny$()) As String()
Dim Mthln: For Each Mthln In Itr(Mthlny)
    PushI MthPmAy, BetBkt(Mthln)
Next
End Function

Function ShtMthPm$(MthPm)
Dim ArgStry$(): ArgStry = SyzLTrim(SplitComma(MthPm))
Dim O$()
Dim Arg: For Each Arg In Itr(ArgStry)
    PushI O, ShtArgzS(Arg)
Next
ShtMthPm = JnSpc(O)
End Function
