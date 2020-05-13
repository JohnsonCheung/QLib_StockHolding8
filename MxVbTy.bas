Attribute VB_Name = "MxVbTy"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbTy."

Function VbTy(TypeName) As VbVarType
Dim O As VbVarType
Select Case RmvSfx(TypeName, "()")
Case "Integer":  O = vbInteger
Case "Boolean":  O = vbBoolean
Case "Byte":     O = vbByte
Case "Currency": O = vbCurrency
Case "Date":     O = vbDate
Case "Decimal": O = vbDecimal
Case "Double":  O = vbDouble
Case "Empty":   O = vbEmpty
Case "Error":   O = vbError
Case "Integer": O = vbInteger
Case "Long":    O = vbLong
Case "Null":    O = vbNull
Case "Object":  O = vbObject
Case "Single":  O = vbSingle
Case "String":  O = vbString
Case "Variant": O = vbVariant
End Select
If HasSfx(TypeName, "()") Then O = O + vbArray
VbTy = O
End Function

Function VbTyNy(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI VbTyNy, TypeName(I)
Next
End Function

Function VbTyAy(Ay) As VbVarType()
Dim V: For Each V In Ay
    PushI VbTyAy, VarType(V)
Next
End Function

Function VbTyAyzTyNy(VbTyNy$()) As VbVarType()
Dim T: For Each T In Itr(VbTyNy)
    PushI VbTyAyzTyNy, VbTy(T)
Next
End Function

Function CvVal(S$, T As VbVarType)
Dim O
Select Case T
Case VbVarType.vbBoolean:   O = CBool(S)
Case VbVarType.vbByte:      O = CByte(S)
Case VbVarType.vbCurrency:  O = CCur(S)
Case VbVarType.vbDate:      O = CDate(S)
Case VbVarType.vbDecimal:   O = CDec(S)
Case VbVarType.vbDouble:    O = CDec(S)
Case VbVarType.vbEmpty:     O = Empty
Case VbVarType.vbError:     O = CVErr(0)
Case VbVarType.vbInteger:   O = CInt(S)
Case VbVarType.vbLong:      O = CLng(S)
Case VbVarType.vbNull:      O = Null
Case VbVarType.vbSingle:    O = CSng(S)
Case VbVarType.vbString:    O = S
Case VbVarType.vbVariant:   O = S
Case Else:
End Select
End Function
