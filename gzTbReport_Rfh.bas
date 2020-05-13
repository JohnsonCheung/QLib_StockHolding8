Attribute VB_Name = "gzTbReport_Rfh"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzTbReport_Rfh."
Option Base 0
Private Const c_FSPec_MB52$ = "MB52 20??-??-??.XLSX"
Sub TbReport_Rfh()
'Aim: Detect any new ZM01/MB52 Fx in import-YYYY folder and add record to table-Report (Y M D ..)
'     When add record to Report: Fields with value is: Y M D.
'     Assume DteCrt has default value Now()
'     Return true if record is added to table Report so that the form caller can Me.Requiry
Dim Ymd() As Ymd: Ymd = MB52Ymd
Dim N%
Dim J%, M As Ymd: For J = 0 To YmdUB(Ymd)
    If InsReport(Ymd(J)) Then
        N = N + 1
    End If
Next
If N > 0 Then
    Sts "[" & N & "] is inserted in Tbl-Report"
End If
If IsFrmOpn("Rpt") Then Form_Rpt.Requery
End Sub
Private Function InsReport(A As Ymd) As Boolean
'Ins to Tbl-Report if @A is new
With CurrentDb.OpenRecordset("Select YY,MM,DD from Report" & OHYmdBexp(A))
    If .EOF Then
        .AddNew
        !YY = A.Y
        !MM = A.M
        !DD = A.D
        .Update
        Sts "Inserting Tbl-Report: " & YYmdStr(A)
    End If
End With
End Function
Private Function MB52Ymd() As Ymd()
Dim Fn$(): Fn = FnAy(MB52IPthPm, c_FSPec_MB52)
Dim I: For Each I In Itr(Fn)
    PushYmd MB52Ymd, YmdzMB52Fn(I)
Next
End Function
Private Function YmdzMB52Fn(Fn) As Ymd
If IsMB52Fn(Fn) Then
    With YmdzMB52Fn
        .Y = Mid(Fn, 8, 2)
        .M = Mid(Fn, 11, 2)
        .D = Mid(Fn, 14, 2)
    End With
End If
End Function
Private Function IsMB52Fn(Fn) As Boolean
Select Case True
Case Not HasMBPfx(Fn), Not HasMBDte(Fn), Not IsXlsx(Fn): Exit Function
End Select
IsMB52Fn = True
End Function
Private Function HasMBPfx(Fn) As Boolean
HasMBPfx = HasPfx(Fn, "MB52 20")
End Function
Private Function HasMBDte(Fn) As Boolean
HasMBDte = IsDate(Mid(Fn, 6, 10))
End Function
