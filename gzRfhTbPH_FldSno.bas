Attribute VB_Name = "gzRfhTbPH_FldSno"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRfhTbPH_FldSno."

Sub RfhTbPH_FldSno()
RfhTbPH_Fld_WithOHxxx
UpdL1
ForEachL1_UpdEachL2
ForEachL2_UpdEachL3
ForEachL3_UpdEachL4
RunCQ "Update ProdHierarchy set Sno=Null,Srt=Null where Not WithOHHst"
End Sub
Private Sub UpdL1()
Dim Sno%
With CurrentDb.OpenRecordset("Select Sno from ProdHierarchy where WithOHHst and Lvl=1 order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub
Private Sub ForEachL1_UpdEachL2()
With CurrentDb.OpenRecordset("Select PH from ProdHierarchy where  WithOHHst and Lvl=1")
    While Not .EOF
        UpdL2 CStr(!PH)
        .MoveNext
    Wend
End With
End Sub
Private Sub UpdL2(L1$)
Dim Sno%
With CurrentDb.OpenRecordset("Select Sno from ProdHierarchy where WithOHHst and Lvl=2 and Left(PH,2)='" & L1 & "' order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub
'
Private Sub ForEachL2_UpdEachL3()
With CurrentDb.OpenRecordset("Select PH from ProdHierarchy where WithOHHst and Lvl=2")
    While Not .EOF
        UpdL3 CStr(!PH)
        .MoveNext
    Wend
End With
End Sub
Private Sub UpdL3(L2$)
Dim Sno%
With CurrentDb.OpenRecordset("Select Sno from ProdHierarchy where WithOHHst and Lvl=3 and Left(PH,4)='" & L2 & "' order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub
'
Private Sub ForEachL3_UpdEachL4()
With CurrentDb.OpenRecordset("Select PH from ProdHierarchy where  WithOHHst and Lvl=3")
    While Not .EOF
        UpdL4 CStr(!PH)
        .MoveNext
    Wend
End With
End Sub
Private Sub UpdL4(L3$)
Dim Sno%
With CurrentDb.OpenRecordset("Select Sno from ProdHierarchy where  WithOHHst and Lvl=4 and Left(PH,7)='" & L3 & "' order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub
