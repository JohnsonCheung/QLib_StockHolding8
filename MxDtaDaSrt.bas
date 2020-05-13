Attribute VB_Name = "MxDtaDaSrt"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaSrt."

Private Sub SrtDrs__Tst()
Dim Drs As Drs, Act As Drs, Ept As Drs, SrtByFF$
GoSub T0
Exit Sub
T0:
    SrtByFF = "A B"
    Drs = DrszFF("A B C", DyoSSVbl("4 5 6|1 2 3|2 3 4"))
    Ept = DrszFF("A B C", DyoSSVbl("1 2 3|2 3 4|4 5 6"))
    GoTo Tst
Tst:
    Act = SrtDrs(Drs, SrtByFF)
    If Not IsEqDrs(Act, Ept) Then Stop
    Return
End Sub

Function SrtDrs(D As Drs, Optional ByDashFF$ = "") As Drs
'@ByDashFF :DashFF ! If @ByDashFF is blank all col of @D is used to sort.  Fld wi dash means descending
'Ret         : ! Srted Drs @@
':DashFF: :SS ! Each Term is a Nm or Dash-Nm
If NoReczDrs(D) Then SrtDrs = D: Exit Function
Dim K As DySrtKey:      K = DySrtKey(ByDashFF, D.Fny)
Dim Dy():               Dy = SrtDy(D.Dy, K)
                    SrtDrs = Drs(D.Fny, Dy)
End Function

Function SrtDy(Dy(), K As DySrtKey) As Variant()
SrtDy = AwIxy(Dy, RxyzSrtDy(SelDy(Dy, K.Cxy), K.IsDes))
End Function

Function SrtDt(A As Dt, Optional SrtByFF$ = "") As Dt
SrtDt = DtzDrs(SrtDrs(DrszDt(A), SrtByFF), A.DtNm)
End Function

Function SrtDyzC(Dy(), C&, Optional IsDes As Boolean) As Variant()
SrtDyzC = SrtDy(Dy, SrtgDySngKey(C, IsDes))
End Function
