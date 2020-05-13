Attribute VB_Name = "gzFc"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzFc."
'1-Good-MH-FcFxFn return 1-StmYM
'1-Good-UD-FcFxFn return 2-StmYM
'Fc 20??-?? MH 8600.xlsx
'Fc 20??-?? MH 8700.xlsx
'Fc 20??-?? UD.xlsx
'         1
'123456789012345678
Private Sub StmYMAyzPth__Tst()
Dim A() As StmYM: A = StmYmAyzPth(FcIPth)
Stop
End Sub

Function StmYmAyzPth(Pth$) As StmYM()
Dim UD$(): UD = FnAy(Pth, "Fc 20??-?? MH.xlsx")
Dim MH$(): MH = FnAy(Pth, "Fc 20??-?? UD.xlsx")
Dim Fn$(): Fn = AddSy(UD, MH)
Dim IFn: For Each IFn In Itr(Fn)
    PushStmYM StmYmAyzPth, StmYmzFcFn(IFn)
Next
End Function
Private Function StmYmzFcFn(FcFn) As StmYM
Dim Y As Byte, M As Byte, Stm$
Stm = StmzFcFn(FcFn): If Stm = "" Then Exit Function
Y = YzFcFn(FcFn): If Y = 0 Then Exit Function
M = MzFcFn(FcFn): If M = 0 Then Exit Function
StmYmzFcFn = StmYM(Stm, Y, M)
End Function

Private Function YzFcFn(FcFn) As Byte
Dim A$: A = Mid(FcFn, 6, 2)
If IsNumeric(A) Then YzFcFn = A
End Function

Private Function StmzFcFn$(FcFn)
StmzFcFn = StmzStm2(Mid(FcFn, 12, 2))
End Function
Private Function MzFcFn(FcFn) As Byte
Dim A%: A = Val(Mid(FcFn, 9, 2))
If 1 <= A And A <= 12 Then MzFcFn = A
End Function
Private Function IsUD(Fn) As Boolean: IsUD = Stm2(Fn) = "UD": End Function
Private Function IsMH(Fn) As Boolean: IsMH = Stm2(Fn) = "MH": End Function
Private Function Stm2$(Fn)
Stm2 = Mid(Fn, 12, 2)
End Function
Private Function IsFc(Fn) As Boolean
Select Case True
Case _
    Left(Fn, 5) <> "Fc 20", _
    Right(Fn, 5) <> ".xlsx", _
    Mid(Fn, 8, 1) <> "-"
Case Else
    IsFc = True
End Select
End Function


Function FcIPth$()
Static P$
If P = "" Then
    P = AppIPth & "Forecast\"
    EnsPth P
End If
FcIPth = P
End Function

Function FcIFxFn$(A As StmYM)
With A
FcIFxFn = FmtQQ("Fc ? ?.xlsx", YymStr(.Y, .M), Stm2zStm(.Stm))
End With
End Function

Function FcIFx$(A As StmYM)
FcIFx = FcIPth & FcIFxFn(A)
End Function

Function FcOPth$()
FcOPth = AppOPth
End Function

Function IsCnlNoFc(A As YM) As Boolean
Dim MH As Boolean, UD As Boolean
MH = HasMhFc(A)
UD = HasUdFc(A)
If MH And UD Then Exit Function
Dim M$
    Dim O$()
    If Not MH Then PushS O, "No MH forecast"
    If Not UD Then PushS O, "No UD forecast"
    M = JnCrLf(O)
    
IsCnlNoFc = MsgBox(M & vbCrLf & vbCrLf & "[Ok]=Continue to generate report" & vbCrLf & "or [Cancel]", vbQuestion + vbOKCancel) = vbCancel
End Function

Private Function HasUdFc(A As YM) As Boolean
HasUdFc = HasRecCQ("Select Top 1 VerYY  from FcSku" & WhereFcStm(StmYM("U", A.Y, A.M)))
End Function

Private Function HasMhFc(A As YM) As Boolean
HasMhFc = HasRecCQ("Select Top 1 VerYY  from FcSku" & WhereFcStm(StmYM("M", A.Y, A.M)))
End Function

Sub FcReadMe()
DoCmd.OpenForm "FcReadMe"
End Sub

Sub ClrFc(A As StmYM, Optional Frm As Access.Form)
Dim Sql$: Sql = "Delete * from Fc" & WhereFcStm(A)
RunCQ Sql
RfhTbFc_FmFcIPth
If Not IsNothing(Frm) Then Frm.Requery
End Sub
