Attribute VB_Name = "JMxFxPrp"
Option Compare Text
Const CMod$ = CLib & "JMxFxPrp."
#If False Then
Option Explicit
Function Wny(Wb As Workbook)
Wny = Itn(Wb.Sheets)
End Function
Function WnyzFx(Fx$, Optional InlAllOtherTbl As Boolean) As String()
ChkFfnExist Fx, "Excel file"
Dim Tny$(), T
Tny = TnyzCat(CatzFx(Fx))
If InlAllOtherTbl Then
    WnyzFx = Tny
    Exit Function
End If
For Each T In Itr(Tny)
    PushNB WnyzFx, WsnzCattn(T)
Next
End Function

Function NoFxw(Fx$, W) As Boolean
NoFxw = Not HasFxw(Fx, W)
End Function

Function HasFxw(Fx$, W) As Boolean
HasFxw = HasEle(WnyzFx(Fx), W)
End Function

Private Function WsnzCattn$(Cattn)
':Cattn: :TblNm ! #Cat-Tbl-Nm#
If HasSfx(Cattn, "FilterDatabase") Then Exit Function
WsnzCattn = RmvSfx(RmvSngQuo(Cattn), "$")
End Function

#End If
