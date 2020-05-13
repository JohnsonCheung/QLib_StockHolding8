Attribute VB_Name = "JMxVbStrQuo"
Option Compare Text
Const CMod$ = CLib & "JMxVbStrQuo."
#If False Then
Option Explicit

Function Quo$(S, C$): Quo = C & S & C: End Function
Function QuoSq$(S): QuoSq = "[" & S & "]": End Function
Function QuoSng$(S): QuoSng = Quo(S, "'"): End Function
Function QuoDbl$(S): QuoDbl = vbDblQ & S & vbDblQ: End Function
Function QuoDte$(S): QuoDte = "#" & S & "#": End Function

Function IsQuoted(S, Q1$, Optional ByVal Q2$) As Boolean
If Q2 = "" Then Q2 = Q1
If FstChr(S) <> Q1 Then Exit Function
IsQuoted = LasChr(S) = Q2
End Function

Function IsSngQuoted(S) As Boolean
IsSngQuoted = IsQuoted(S, "'")
End Function
#End If
