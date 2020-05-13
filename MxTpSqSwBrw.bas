Attribute VB_Name = "MxTpSqSwBrw"
Option Explicit
Option Compare Text
Private Function FmtSqSw(A As SqSw) As String()
ClrXX
X FmtDic(A.FldSw.Pm, Tit:="Pm")
X FmtDic(A.StmtSw, Tit:="StmtSw")
X FmtDic(A.FldSw, Tit:="FldSw")
FmtSqSw = XX
End Function
Private Sub BrwSwly(L() As Swl): BrwAy FmtSwly(L): End Sub
Function FmtSwly(L() As Swl) As String()
Dim J%: For J = 0 To SwlUB(L)
    PushI FmtSwly, FmtSwl(L(J))
Next
End Function
Function FmtSwl$(L As Swl)
With L
Dim X$
Select Case True
Case IsOrAndStr(.Op): X = JnSpc(.Termy)
Case IsEqNeStr(.Op): X = .Termy(0) & " " & .Termy(1)
End Select
FmtSwl = JnSpcAp(.Swn, CvBoolOp(.Op), X)
End With
End Function
