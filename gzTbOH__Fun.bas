Attribute VB_Name = "gzTbOH__Fun"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzTbOH__Fun."


Function GitYpStk%()
GitYpStk = CurrentDb.OpenRecordset("Select YpStk from YpStk where NmYpStk='GIT'").Fields(0).Value
End Function
