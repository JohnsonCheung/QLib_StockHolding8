Attribute VB_Name = "MxDaoDefFd"
Option Compare Text
Option Explicit
Const CNs$ = "Def"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDefFd."

Function CvFd(A) As DAO.Field
Set CvFd = A
End Function

Function CvFd2(A) As DAO.Field2
Set CvFd2 = A
End Function

Function CloneFd(A As DAO.Field2, Fldn) As DAO.Field2
Set CloneFd = New DAO.Field
With CloneFd
    .Name = Fldn
    .Type = A.Type
    .AlloZZeroLength = A.AlloZZeroLength
    .Attributes = A.Attributes
    .DefaultValue = A.DefaultValue
    .Expression = A.Expression
    .Required = A.Required
    .ValidationRule = A.ValidationRule
    .ValidationText = A.ValidationText
End With
End Function

Function IsEqFd(A As DAO.Field2, B As DAO.Field2) As Boolean
With A
    If .Name <> B.Name Then Exit Function
    If .Type <> B.Type Then Exit Function
    If .Required <> B.Required Then Exit Function
    If .AlloZZeroLength <> B.AlloZZeroLength Then Exit Function
    If .DefaultValue <> B.DefaultValue Then Exit Function
    If .ValidationRule <> B.ValidationRule Then Exit Function
    If .ValidationText <> B.ValidationText Then Exit Function
    If .Expression <> B.Expression Then Exit Function
    If .Attributes <> B.Attributes Then Exit Function
    If .Size <> B.Size Then Exit Function
End With
IsEqFd = True
End Function

Function Fdv(A As DAO.Field)
On Error Resume Next
Fdv = A.Value
End Function
