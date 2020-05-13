Attribute VB_Name = "gzPH__Fun"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzPH__Fun."
Function PH_Where$(Lvl As Byte, WithOHCur As Boolean, WithOHHst As Boolean)
Dim L$
Select Case Lvl
Case 1: L = "Lvl=1"
Case 2: L = "Lvl<=2"
Case 3: L = "Lvl<=3"
Case Else: L = "True"
End Select
Dim WithOH$
Select Case True
Case WithOHCur: WithOH = " and WithOHCur"
Case WithOHHst: WithOH = " and WithOHHst"
End Select
PH_Where = " Where " & L & WithOH
End Function


Function PH7Jn() As String()
Dim O$(6)
O(0) = "x.Stm=a.Stm"
O(1) = "x.Stm=a.Stm and x.BusArea=a.BusArea"
O(2) = "x.Stm=a.Stm and x.PHL1=a.PHL1"
O(3) = "x.Stm=a.Stm and x.PHL2=a.PHL2"
O(4) = "x.Stm=a.Stm and x.PHL3=a.PHL3"
O(5) = "x.Stm=a.Stm and x.PHL4=a.PHL4"
O(6) = "x.Sku=a.Sku"
PH7Jn = O
End Function
Function PH7Key(Optional Alias$) As String()
Dim O$(6)
Dim A$: A = Alias
O(0) = RplQ("?Stm", A)
O(1) = RplQ("?Stm,?BusArea", A)
O(2) = RplQ("?Stm,?PHL1", A)
O(3) = RplQ("?Stm,?PHL2", A)
O(4) = RplQ("?Stm,?PHL3", A)
O(5) = RplQ("?Stm,?PHL4", A)
O(6) = RplQ("?Sku", A)
PH7Key = O
End Function
