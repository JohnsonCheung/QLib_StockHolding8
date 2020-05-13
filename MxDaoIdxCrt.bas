Attribute VB_Name = "MxDaoIdxCrt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoIdxCrt."
Type TFny
    Tbn As String
    Fny() As String
End Type
Function TFny(Tbn, Fny$()) As TFny
With TFny
    .Tbn = Tbn
    .Fny = Fny
End With
End Function
Function TFnySi&(A() As TFny): On Error Resume Next: TFnySi = UBound(A) + 1: End Function
Function TFnyUB&(A() As TFny): TFnyUB = TFnySi(A) - 1: End Function
Sub PushTFny(O() As TFny, M As TFny): Dim N&: N = TFnySi(O): ReDim O(N): O(N) = M: End Sub
Function TFnyzFF(T, FF$) As TFny
TFnyzFF = TFny(T, FnyzFF(FF))
End Function

Sub CrtPk(D As Database, T)
D.Execute sqlCrtPk(T)
End Sub

Sub CrtSk(D As Database, T, Skff$)
D.Execute sqlCrtSkzFF(T, Skff)
End Sub

Sub CrtUKey(D As Database, T, K$, FF$)
D.Execute sqlCrtUKy(T, K, FF)
End Sub

Sub CrtKey(D As Database, T, K$, FF$)
D.Execute sqlCrtKey(T, K, FF)
End Sub

Function sqlCrtPk$(T)
sqlCrtPk = FmtQQ("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function

Function SqyCrtPk(Tny$()) As String()
Dim T: For Each T In Itr(Tny)
    PushI SqyCrtPk, sqlCrtPk(T)
Next
End Function

Function sqlCrtSk$(A As TFny)
sqlCrtSk = FmtQQ("Create unique Index SecondaryKey on [?] (?)", A.Tbn, JnComma(QuoTerm(A.Fny)))
End Function

Function SqyCrtSk(A() As TFny) As String()
Dim J%: For J = 0 To TFnyUB(A)
    PushI SqyCrtSk, sqlCrtSk(A(J))
Next
End Function

Function sqlCrtSkzFF$(T, Skff$)
sqlCrtSkzFF = sqlCrtSk(TFnyzFF(T, Skff))
End Function

Function sqlCrtUKy$(T, K$, FF$)

End Function

Function sqlCrtKey$(T, K$, FF$)

End Function
