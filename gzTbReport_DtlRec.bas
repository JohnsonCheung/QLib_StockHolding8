Attribute VB_Name = "gzTbReport_DtlRec"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzTbReport_DtlRec."
Sub DltMB52Rec(A As Ymd)
With A
If Not Start("Delete for [" & .Y + 2000 & "-" & Format(.M, "00") & "-" & Format(.D, "00") & "]?", "Delete?") Then Exit Sub
End With
RunCQ "Delete from Report" & OHYmdBexp(A)
End Sub
