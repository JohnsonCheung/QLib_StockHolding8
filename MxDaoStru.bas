Attribute VB_Name = "MxDaoStru"
Option Compare Text
Option Explicit
Const CNs$ = "sdfsdf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoStru."

Sub DmpStru(D As Database)
Dmp Stru(D)
End Sub

Sub DmpStruTT(D As Database, TT$)
Dmp StruTT(D, TT)
End Sub

Function StruFld(ParamArray Ap()) As Drs
Dim Dy(), S$, I, Av(), Ele$, LikFF$, LikFld$, J
Av = Ap
For Each I In Av
    S = I
    AsgTRst S, Ele, LikFF
    For Each J In SyzSS(LikFF)
        LikFld = J
        PushI Dy, Array(Ele, LikFld)
    Next
Next
StruFld = DrszFF("Ele FldLik", Dy)
End Function

Function StruInf(D As Database) As Dt
Dim T$, TT, Dy(), Des$, NRec&, Stru$
'For Each TT In TnyDb(D)
    T = TT
'    Des = Dbt_Des(D, T)
'    Stru = RmvT1(Stru(D, T))
'    NRec = NRecDT(D, T)
    PushI Dy, Array(T, NRec, Des, Stru)
'Next
StruInf = DtByFF("Tbl", "Tbl NRec Des", Dy)
End Function

'**Stru
Function StruD() As String(): StruD = Stru(CDb): End Function
Function Stru(D As Database) As String(): Stru = AliLyz1T(StruTny(D, Tny(D))): End Function
Function StruTny(D As Database, Tny$()) As String()
Dim I: For Each I In Itr(QSrt(Tny))
    PushI StruTny, StruT(D, I)
Next
End Function
Function StruzRs$(A As DAO.Recordset)
Dim O$(), F As DAO.Field2
For Each F In A.Fields
    PushI O, FdStr(F)
Next
StruzRs = JnCrLf(O)
End Function
Function StruCT$(T): StruCT = StruT(CDb, T): End Function
Function StruCTT(TT$) As String(): StruCTT = StruTT(CDb, TT): End Function
Function StruT$(D As Database, T)
Dim F$()
    F = Fny(D, T)
    F = AmRplStar(F, T)
    F = QuoTermy(F)

Dim Pk$()
    Pk = PkFny(D, T)
    Pk = AmRplStar(Pk, T)
    Pk = QuoTermy(Pk)
    
Dim P$
    P = JnSpc(Pk)
    If P <> "" Then P = " " & P & " |"

Dim R$
    Dim Rst$()
    Rst = MinusAy(F, Pk)
    R = " " & JnSpc(QuoTermy(Rst))
StruT = T & P & R
End Function
Function StruTT(D As Database, TT$): StruTT = StruTny(D, Ny(TT)): End Function
