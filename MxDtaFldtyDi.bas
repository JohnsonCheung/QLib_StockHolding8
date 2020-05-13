Attribute VB_Name = "MxDtaFldtyDi"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CNs$ = "Stru"
Const CMod$ = CLib & "MxDtaFldtyDi."
Function DiFqShtDao(FxOrFb$, TorW) As Dictionary
Const CSub$ = CMod & "DiFqShtDao"
Select Case True
Case IsFb(FxOrFb): Set DiFqShtDao = DiFqShtDaoTyzFbt(FxOrFb, TorW)
Case IsFx(FxOrFb): Set DiFqShtDao = DiFqShtAdozFxw(FxOrFb, TorW)
Case Else: Thw CSub, "FxOrFb should be Fx or Fb", "FxOrFb TorW", FxOrFb, TorW
End Select
End Function

Function DiFqShtDaoTyzFbt(Fb, T) As Dictionary
Set DiFqShtDaoTyzFbt = New Dictionary
Dim D As Database: Set D = Db(Fb)
Dim Td As DAO.TableDef: Set Td = D.TableDefs(T)
Dim F As DAO.Field: For Each F In Td.Fields
    DiFqShtDaoTyzFbt.Add F.Name, ShtDaoTy(F.Type)
Next
End Function

Function DiFqShtAdozFxw(Fx, Optional W = "Sheet1") As Dictionary
Set DiFqShtAdozFxw = New Dictionary
Dim Cat As Catalog: Set Cat = CatzFx(Fx)
Dim C As Adox.Column: For Each C In Cat.Tables(AxTbn(W)).Columns
    DiFqShtAdozFxw.Add C.Name, ShtAdoTy(C.Type)
Next
End Function

Private Sub DiFqShtAdozFxw__Tst()
BrwDic DiFqShtAdozFxw(SalTxtFx)
End Sub
