Attribute VB_Name = "MxXlsWsPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsWsPrp."

Function HasLo(S As Worksheet, Lon$) As Boolean
HasLo = HasItn(S.ListObjects, Lon)
End Function

Function LasCell(S As Worksheet) As Range
Set LasCell = S.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function LasRno&(S As Worksheet)
LasRno = LasCell(S).Row
End Function

Function LasCno%(S As Worksheet)
LasCno = LasCell(S).Column
End Function

Function PtNyzWs(S As Worksheet) As String()
PtNyzWs = Itn(S.PivotTables)
End Function

Property Get MaxCnozX&(X As Excel.Application)
MaxCnozX = IIf(X.Version = "16.0", 16384, 255)
End Property

Property Get MaxRnozX&(X As Excel.Application)
MaxRnozX = IIf(Xls.Version = "16.0", 1048576, 65535)
End Property
