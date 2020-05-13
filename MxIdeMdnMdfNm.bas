Attribute VB_Name = "MxIdeMdnMdfNm"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdnMdfNm."
Type MdfNm
    IsPrv As Boolean
    Nm As String
End Type
Function MdfNm(IsPrv As Boolean, Nm$) As MdfNm
With MdfNm
    .IsPrv = IsPrv
    .Nm = Nm
End With
End Function
