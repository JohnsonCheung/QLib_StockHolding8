Attribute VB_Name = "MxVbStrJnQuo"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbStrJnQuo."


Function JnQDblComma$(Sy$())
JnQDblComma = JnComma(AmQuoDbl(Sy))
End Function

Function JnQDblSpc$(Sy$())
JnQDblSpc = JnSpc(AmQuoDbl(Sy))
End Function

Function JnQSngComma$(Sy$())
JnQSngComma = JnComma(AmQuoSng(Sy))
End Function

Function JnQSngSpc$(Sy$())
JnQSngSpc = JnSpc(AmQuoSng(Sy))
End Function

Function JnQSqCommaSpc$(Sy$())
JnQSqCommaSpc = JnCommaSpc(AmQuoSq(Sy))
End Function

Function JnQSqBktSpc$(Ay)
JnQSqBktSpc = JnSpc(AmQuoSq(Ay))
End Function
