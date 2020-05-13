Attribute VB_Name = "MxXlsRgCellPrp"
Option Explicit
Option Compare Text
Const CNs$ = "Cell.Prp"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRgCellPrp."

Function CellBelow(Cell As Range, Optional N = 1) As Range
Set CellBelow = RgRC(Cell, 1 + N, 1)
End Function

Function CellAbove(Cell As Range, Optional Above = 1) As Range
Set CellAbove = RgRC(Cell, 1 - Above, 1)
End Function

Function CellRight(A As Range, Optional Right = 1) As Range
Set CellRight = RgRC(A, 1, 1 + Right)
End Function
