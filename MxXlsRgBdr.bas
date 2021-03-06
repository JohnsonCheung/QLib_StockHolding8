Attribute VB_Name = "MxXlsRgBdr"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CNs$ = "Fmt.Bdr"
Const CMod$ = CLib & "MxXlsRgBdr."
Sub BdrBottom(A As Range)
Bdr A, xlEdgeBottom
If A.Row < MaxRno Then
    Bdr RgR(A, NRoZZRg(A) + 1), xlEdgeTop
End If
End Sub

Sub BdrInside(A As Range)
Bdr A, xlInsideHorizontal
Bdr A, xlInsideVertical
End Sub

Sub BdrLeft(A As Range)
Bdr A, xlEdgeLeft
If A.Column > 0 Then
    Bdr RgC(A, NColzRg(A) - 1), xlEdgeRight
End If
End Sub

Sub BdrRight(A As Range)
Bdr A, xlEdgeRight
If A.Column < MaxCno Then
    Bdr RgC(A, NColzRg(A) + 1), xlEdgeLeft
End If
End Sub

Sub BdrTop(A As Range)
Bdr A, xlEdgeTop
If A.Row > 1 Then
    Bdr RgR(A, 0), xlEdgeBottom
End If
End Sub

Sub Bdr(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub
Sub BdrNone(A As Range, Ix As XlBordersIndex)
A.Borders(Ix).LineStyle = xlLineStyleNone
End Sub

Sub BdrAround(A As Range)
Exit Sub
BdrLeft A
BdrRight A
BdrTop A
BdrBottom A
End Sub

Sub BdrAroundNone(A As Range)
BdrNone A, xlInsideHorizontal
BdrNone A, xlInsideVertical
BdrNone A, xlEdgeLeft
BdrNone A, xlEdgeRight
BdrNone A, xlEdgeBottom
BdrNone A, xlEdgeTop
End Sub

Sub BdrRg(R As Range, BdrIx As XlBordersIndex)
Select Case BdrIx
Case XlBordersIndex.xlEdgeLeft
Case XlBordersIndex.xlEdgeRight
Case Else

End Select
With R.Borders(BdrIx)
    .LineStyle = XlLineStyle.xlContinuous
    .Weight = XlBorderWeight.xlMedium
End With
End Sub
Sub BdrRgLeft(R As Range): BdrRg R, xlEdgeLeft: End Sub
Sub BdrRgRight(R As Range): BdrRg R, xlEdgeRight: End Sub

Sub BdrRgAy(A() As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
Dim I
For Each I In Itr(A)
    Bdr CvRg(I), Ix, Wgt
Next
End Sub

Function BdrLoAround(A As ListObject)
Dim R As Range
Set R = RgMoreTop(A.DataBodyRange)
If A.ShowTotals Then Set R = RgMoreBelow(R)
BdrAround R
End Function
