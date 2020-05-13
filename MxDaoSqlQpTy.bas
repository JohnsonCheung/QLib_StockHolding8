Attribute VB_Name = "MxDaoSqlQpTy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDaoSqlQpTy."
Function sqlTyzDyC$(Dy(), C&)
sqlTyzDyC = sqlTyzAv(ColzDy(Dy, C))
End Function
Function sqlTyzAv$(Av())
Dim O As VbVarType, V, T As VbVarType
For Each V In Av
    T = VarType(V)
    If T = vbString Then
        If Len(V) > 255 Then sqlTyzAv = "Memo": Exit Function
    End If
'    O = MaxVbTy(O, T)
Next
End Function
Function sqlTyzVbTy$(Dy As VbVarType)
Dim O$
Select Case Dy
Case vbEmpty:   O = "Text(255)"
Case vbBoolean: O = "YesNo"
Case vbByte:    O = "Byte"
Case vbInteger: O = "Short"
Case vbLong:    O = "Long"
Case vbDouble:  O = "Double"
Case vbSingle:  O = "Single"
Case vbCurrency: O = "Currency"
Case vbDate:    O = "Date"
Case vbString:  O = "Text(255)"
Case Else: Stop
End Select
sqlTyzVbTy = O
End Function
