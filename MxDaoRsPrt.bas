Attribute VB_Name = "MxDaoRsPrt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoRsPrt."

Function CslzRsFny$(A As DAO.Recordset, Fny$())
CslzRsFny = CslzFdsFny(A.Fields, Fny)
End Function

Function CslzRs$(A As DAO.Recordset)
CslzRs = CslzFds(A.Fields)
End Function

Function CsyzRs(A As DAO.Recordset, Optional FF$) As String()
Dim Fny$(): Fny = FnyzFFRs(FF, A)
A.MoveFirst
While Not A.EOF
    PushI CsyzRs, CslzRsFny(A, Fny)
    A.MoveNext
Wend
End Function

Function CsvzRs1(A As DAO.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As DAO.Field
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "CslzRsLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    PushI CsvzRs1, CslzRs(A)
    A.MoveNext
Wend
End Function

Function JnRs(A As DAO.Recordset, Optional Sep$ = " ", Optional FF$) As String()
Dim O$()
With A
    Push O, Join(FnyzRs(A), Sep)
    While Not .EOF
        Push JnRs, RsLin(A, Sep)
        .MoveNext
    Wend
End With
JnRs = O
End Function

Function FmtRec(A As DAO.Recordset, Optional FF$) As String()
Dim Fny$(): Fny = FnyzFFRs(FF, A)
FmtRec = FmtNyAv(Fny, DrzRsFny(A, Fny))
End Function

Function FnyzFFRs(FF$, A As DAO.Recordset) As String()
If FF = "" Then
    FnyzFFRs = FnyzRs(A)
Else
    FnyzFFRs = FnyzFF(FF)
End If
End Function
