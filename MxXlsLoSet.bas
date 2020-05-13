Attribute VB_Name = "MxXlsLoSet"
Option Explicit
Option Compare Text
Const CNs$ = "Set.Lo"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLoSet."

Sub SetLoAutoFit(L As ListObject, Optional MaxW = 100)
Dim C As Range: Set C = LoAllEntCol(L)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = EntRgC(C, J)
   If EntC.ColumnWidth > MaxW Then EntC.ColumnWidth = MaxW
Next
End Sub
Sub SetLon(L As ListObject, Lon$)
Const CSub$ = CMod & "Fmtn"
If Lon <> "" Then
    If Not HasLo(WszLo(L), Lon) Then
        L.Name = Lon
    Else
        Inf CSub, "Lo"
    End If
End If
End Sub
