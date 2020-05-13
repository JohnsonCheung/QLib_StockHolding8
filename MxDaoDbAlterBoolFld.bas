Attribute VB_Name = "MxDaoDbAlterBoolFld"
Option Compare Text
Const CMod$ = CLib & "MxDaoDbAlterBoolFld."
#If False Then
Option Explicit
Sub AltBoolFld(T$, F$, Optional TrueStr$ = "Y", Optional FalseStr$ = "N")
RenFld T, F, F & "(Bool)"
Dim L%: L = Max(Len(TrueStr), Len(FalseStr))
RunCQ FmtQQ("Alter table [?] Add Column [?] Text(?)", T, F, L)
RunCQ FmtQQ("Update [?] set [?]=IIf(IsNull([?]),'',IIf([?(Bool)],'?','?'))", T, F, F, F, TrueStr, FalseStr)
RunCQ FmtQQ("Alter table [?] Drop Column [?(Bool)]", T, F)
End Sub
Sub RenFld(T, F, NewF$)
CurrentDb.TableDefs(T).Fields(F).Name = NewF
End Sub

#End If
