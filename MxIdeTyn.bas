Attribute VB_Name = "MxIdeTyn"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeTyn."

Function IsUdtTyn(Tyn$) As Boolean
Static X$(): If Si(X) = 0 Then X = UdtnyP
IsUdtTyn = HasEle(X, Tyn)
End Function

Function IsEnmTyn(Tyn$) As Boolean
Static X$(): If Si(X) = 0 Then X = EnmnyP
IsEnmTyn = HasEle(X, Tyn)
End Function

Function IsObjTyn(Tyn$) As Boolean ' return true if @Tyn (isBlnk | IsPrimTy | IsUdtTyn | IsEnmTyn)
If Tyn = "" Then Exit Function
If IsPrimTy(Tyn) Then Exit Function
If IsUdtTyn(Tyn) Then Exit Function
If IsEnmTyn(Tyn) Then Exit Function
IsObjTyn = True
End Function

Function ShtTyn$(Tyn$) ' return ShtTyn for some known class
Dim O$
Select Case Tyn
Case "VbProject": O = "Pj"
Case "Access.Application": O = "Acs"
Case "Access.Control": O = "AcsCtl"
Case "Access.CommandButton": O = "AcsBtn"
Case "Access.ToggleButton": O = "AcsTgl"
Case "Excel.Application": O = "Xls"
Case "Excel.Addin": O = "XlsAddin"
Case "Range": O = "Rg"
Case "ListObject": O = "Lo"
Case "ListObject()": O = "LoAy"
Case "Excel.Worksheet", "Worksheet": O = "Ws"
Case "Excel.Workbook", "Workbook": O = "Wb"
Case "ADODB.Recordset": O = "Ars"
Case "ADODB.Connection": O = "Cn"
Case "ADOX.Table": O = "AdoTd"
Case "ADODB.DataTypeEnum": O = "AdoTy"
Case "VBA.Collection", "Collection": O = "Coll"
Case "Dictionary": O = "Dic"
Case "CodeModule": O = "Md"
Case "VBComponent": O = "Cmp"
Case "vbext_ComponentType": O = "eCmpTy"
Case "Database": O = "Db"
Case Else: O = Tyn
End Select
ShtTyn = O
End Function
