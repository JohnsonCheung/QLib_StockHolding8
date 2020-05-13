Attribute VB_Name = "MxXlsWsChkRecMsg"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWsChkRecMsg."

Private Const Col_$ _
                            = "[?] "
Private Const BlnkQQ$ _
                            = Col_ & "has some blank"
Private Const MisQQ$ _
                            = Col_ & "misses [?]"
Private Const NInLisQQ$ _
                            = Col_ & "with value [?] is not one of ?"
Private Const NInTblQQ$ _
                            = Col_ & "with value [?] is not in table [?]"
Function NAllIn(Fx$, W$, ShdAllIn() As FldVy) As String()

End Function

Function NAllInSingle(Fx$, W$, ShdAllIn As FldVy) As String()

End Function

Function Blnk(Fx$, W$, Fny$()) As String()

End Function


Function RecEmsg$(A As RecErDta)
'Select Case True
'Case T = NoEr
''Case T = Mis:   RecEmsg = ColMisMsgl(A.Mis)
''Case T = Blnk:  RecEmsg = ColBlnkMsgl(A.F)
'Case T = NInLis: RecEmsg = ColNInLisMsgl(A.F, A.Val)
'Case T = NInTbl: RecEmsg = ColNInTblMsgl(A.F, A.Val)
'Case T = NInTbc: RecEmsg = ColNInTbcMsgl(A.F, A.Val, A.T)
'Case Else: PmEr CSub, "ColErTy", T, ColErTySS
'End Select
End Function

Private Function ColMisMsgl$(F, V)
ColMisMsgl = FmtQQ(MisQQ, F, V) & vbCrLf
End Function

Private Function ColNInLisMsgl$(F, E, V)
'ColNInLisMsgl = FmtQQ(ColNInLisQQ, F, E, V) & vbCrLf
End Function

Private Function ColNInTbcMsgl$(F, E, T, C)
'ColNInTbcMsgl = FmtQQ(ColNInLisQQ, F, E, C) & vbCrLf
End Function

Private Function ColNInTblMsgl$(F, E, T)
'ColNInTblMsgl = FmtQQ(ColNInTblQQ, F, E, T) & vbCrLf
End Function

Private Function ColBlnkMsgl$(C$)
'ColBlnkMsgl$ = FmtQQ(ColBlnkQQ, C) & vbCrLf
End Function
