Attribute VB_Name = "MxDaoAdoAxtDrs"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "AdoAxtDrs."
Public Const AxtFF$ = "Tbn Name Type DefinedSize NumericScale Precision RelatedColumn SortOrder"
Private Sub AxColDrszFxw__Tst()
BrwDrs AxColDrszFxw(MB52LasIFx, MB52Wsn)
End Sub
Function AxColDrszFxw(Fx, W) As Drs
Dim C As Catalog: Set C = CatzFx(Fx)
Dim T As Table: Set T = C.Tables(AxTbn(W))
AxColDrszFxw = AxColDrs(T)
End Function

Function AxColDrs(T As Adox.Table) As Drs
AxColDrs = DrszFF(AxtFF, AxColDy(T))
End Function

Private Function AxColDy(T As Adox.Table) As Variant()
Dim C As Adox.Column: For Each C In T.Columns
    PushI AxColDy, AxColDr(T.Name, C)
Next
End Function

Private Function AxColDr(Tbn$, C As Adox.Column) As Variant()
With C
AxColDr = Array(Tbn, .Name, .Type, .DefinedSize, .NumericScale, .Precision, RelatedColumn(C), SortOrder(C))
End With
End Function

Private Function RelatedColumn$(C As Adox.Column)
On Error Resume Next
RelatedColumn = C.RelatedColumn
End Function

Private Function SortOrder(C As Adox.Column) As SortOrderEnum
On Error Resume Next
SortOrder = C.SortOrder
End Function
