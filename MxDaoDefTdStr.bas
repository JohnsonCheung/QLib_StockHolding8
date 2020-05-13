Attribute VB_Name = "MxDaoDefTdStr"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoDefTdStr."

Function TdStr$(A As DAO.TableDef)
Dim T$, Id$, S$, R$
    T = A.Name
    If HasPkzTd(A) Then Id = "*Id"
    Dim Pk$(): Pk = Sy(T & "Id")
    Dim Sk$(): Sk = SkFnyzTd(A)
    If HasSkzTd(A) Then S = Tml(AmRpl(Sk, T, "*")) & " |"
    R = Tml(CvSy(MinusAyAp(FnyzTd(A), Pk, Sk)))
TdStr = JnSpc(SyNB(T, Id, S, R))
End Function

Function TdStrzT$(D As Database, T)
TdStrzT = TdStr(D.TableDefs(T))
End Function
