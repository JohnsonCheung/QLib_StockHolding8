Attribute VB_Name = "gzRfhTbPH_FldSrt"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRfhTbPH_FldSrt."
Sub RfhTbPH_FldSrt()
DoCmd.SetWarnings False
Dim D As Dictionary: Set D = DiPHqSno
With CurrentDb.OpenRecordset("SELECT Srt,PH FROM ProdHierarchy where WithOHHst")
    While Not .EOF
        .Edit
        !Srt = FndSrt(D, (Nz(!PH, "")))
        .Update
        .MoveNext
    Wend
End With
End Sub
Private Function DiPHqSno() As Dictionary
Set DiPHqSno = New Dictionary
With CurrentDb.OpenRecordset("Select PH,format(Nz(x.Sno,0),'00') As Sno from ProdHierarchy x")
    While Not .EOF
        DiPHqSno.Add CStr(!PH), CStr(!Sno)
        .MoveNext
    Wend
End With
End Function
Private Function FndSrt$(DiPHqSno As Dictionary, PH$)
Dim D As Dictionary: Set D = DiPHqSno
Select Case Len(PH)
Case 2: FndSrt = D(PH)
Case 4: FndSrt = D(Left(PH, 2)) & D(PH)
Case 7: FndSrt = D(Left(PH, 2)) & D(Left(PH, 4)) & D(PH)
Case 10: FndSrt = D(Left(PH, 2)) & D(Left(PH, 4)) & D(Left(PH, 7)) & D(PH)
End Select
End Function

Private Function FndSno$(DiPHqSno As Dictionary, PH$)
FndSno = CurrentDb.OpenRecordset("Select Format(Sno,'00') from ProdHierarchy where PH='" & PH & "'").Fields(0).Value
End Function

Private Sub Upd_L2()
RunCQ "Update ProdHierarchy set Srt=Format(Sno,'00') where Lvl=2"
End Sub
