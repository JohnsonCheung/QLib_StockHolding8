Attribute VB_Name = "MxIdeSrcCSubEns"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CNs$ = "CSub"
Const CMod$ = CLib & "MxIdeSrcCSubEns."

Sub EnsCSubP(Optional Upd As eUpdRpt)
EnsCSubzP CPj, Upd
End Sub

Sub EnsCSubM(Optional Upd As eUpdRpt)
EnsCSubzM CMd, Upd
End Sub

Sub EnsCSubzP(P As VBProject, Optional Upd As eUpdRpt)
Dim D As Drs: D = MthlDrszP(P)
                  EnsCSubzMthlDrs D, Upd
End Sub

Private Sub EnsCSubzM__Tst()
Dim M As CodeModule
GoSub T1
'GoSub YY
Exit Sub
YY:  Set M = Md("MxEnsMdy"): GoTo Tst
T1:  Set M = Md("MxCmpOp"):   GoTo Tst
Tst: EnsCSubzM M, eUpdAndRpt: Return
End Sub

Private Sub EnsCSubP__Tst()
EnsCSubP eUpdAndRpt
End Sub

Sub EnsCSubzM(M As CodeModule, Optional Upd As eUpdRpt)
EnsCSubzMthlDrs MthlDrszM(M), Upd
End Sub

Private Sub EnsCSubzMthlDrs__Tst()
Dim D As Drs
GoSub YY
Exit Sub
YY:  D = MthlDrszMthn(Md("MxCmpOp"), "SetCmpNm"): GoTo Tst
Tst: EnsCSubzMthlDrs D, eUpdOnly: Return
End Sub

Sub EnsCSubzMthlDrs(MthlDrs As Drs, Upd As eUpdRpt)
Dim D1 As Drs: D1 = CSubDrs(MthlDrs)  ' Mdn L Mthl CurCSubLno CurCSubLin EptCSubLno EptCSubLin OpLno LinOp OldL NewL
Dim D2 As Drs: D2 = SelDrs(D1, MdLnOpFF)
                    MdyLnzUpd D2, Upd              '<==
End Sub
