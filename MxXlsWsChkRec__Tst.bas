Attribute VB_Name = "MxXlsWsChkRec__Tst"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWsChkRec__Tst."

Private Sub ChkWsRec__Tst()
Dim Fx$, W$, RecShdSy$(), RecShdAy() As RecShd
Const P$ = "ChkWsRec__Tst\"
GoSub T1
Exit Sub
T1:
    Fx = TFx1
    W = "Sheet1"
    PushRecShd RecShdAy, RecShdNBlnk("Material")
    
    RecShdSy = Sy("Material ShdNonBlnk")
    OpnFxMax Fx: Stop
    GoTo Tst
T2:
    Fx = TFx2
    W = "Sheet1"
    RecShdSy = Sy()
    OpnFx Fx: Stop
    GoTo Tst
Tst:
    ChkWsRec Fx, W, RecShdSy
    Return
End Sub
Private Sub CrtTFx1(): CpyFfn MB52LasIFx, EnsFfnAllFdr(TFx1): End Sub
Private Sub CrtTFx2(): CpyFfn MB52LasIFx, EnsFfnAllFdr(TFx2): End Sub
Private Function Pseg$(): Pseg = "ChkWsRec__Tst\":        End Function
Private Property Get TFx1$(): TFx1 = ResFfn(Pseg & "Fx1.xlsx"): End Property
Private Property Get TFx2$(): TFx2 = ResFfn(Pseg & "Fx2.xlsx"): End Property
