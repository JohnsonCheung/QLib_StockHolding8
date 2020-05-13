Attribute VB_Name = "MxDtaDte"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dta.Dte"
Const CMod$ = CLib & "MxDtaDte."

Function CurMM() As Byte
CurMM = Month(Now)
End Function

Function FstDte(A As Date) As Date
FstDte = DateSerial(Year(A), Month(A), 1)
End Function

Function IsVdtDte(A) As Boolean
On Error Resume Next
IsVdtDte = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function LasDte(A As Date) As Date
LasDte = PrvDte(FstDte(NxtMthFstDte(A)))
End Function

Function NxtMthFstDte(A As Date) As Date
NxtMthFstDte = FstDte(DateAdd("M", 1, A))
End Function

Function IsHHMMDD(S) As Boolean
Select Case True
Case _
    Len(S) <> 6, _
    Not IsHH(Left(S, 2)), _
    Not Is0059(Mid(S, 3, 2)), _
    Not Is0059(Right(S, 2))
Case Else: IsHHMMDD = True
End Select
End Function

Function IsHH(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "23"
Case Else: IsHH = True
End Select
End Function

Function Is0059(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "59"
Case Else: Is0059 = True
End Select
End Function

Function IsYYYYMMDD(S) As Boolean
If Len(S) <> 8 Then Exit Function
If Not IsYYYY(Left(S, 4)) Then Exit Function
If Not IsMM(Mid(S, 5, 2)) Then Exit Function
If Not IsDD(Right(S, 2)) Then Exit Function
IsYYYYMMDD = True
End Function

Function IsMM(S) As Boolean
If Len(S) <> 2 Then Exit Function
If Not IsAllDig(S) Then Exit Function
If S < "00" Then Exit Function
If S > "12" Then Exit Function
IsMM = True
End Function
Function IsYYYY(S) As Boolean
Select Case True
Case Len(S) <> 4, Not IsAllDig(S), S < "2000"
Case Else: IsYYYY = True
End Select
End Function
Function IsDD(S) As Boolean
Select Case True
Case Len(S) <> 2, Not IsAllDig(S), S < "00", "31" < S
Case Else: IsDD = True
End Select
End Function

Function PrvDte(A As Date) As Date
PrvDte = DateAdd("D", -1, A)
End Function

Function YYMM$(A As Date)
YYMM = Right(Year(A), 2) & Format(Month(A), "00")
End Function

Function FstDtezYYMM(YYMM) As Date
FstDtezYYMM = DateSerial(Left(YYMM, 2), Mid(YYMM, 3, 2), 1)
End Function

Function FstDtezYM(Y As Byte, M As Byte) As Date
FstDtezYM = DateSerial(2000 + Y, M, 1)
End Function

Function LasDtezYM(Y As Byte, M As Byte) As Date
LasDtezYM = NxtMthFstDte(FstDtezYM(Y, M))
End Function

Function YofNxtMzYM(Y As Byte, M As Byte) As Byte
YofNxtMzYM = IIf(M = 12, Y + 1, Y)
End Function

Function YofPrvMzYM(Y As Byte, M As Byte) As Byte
YofPrvMzYM = IIf(M = 1, Y - 1, Y)
End Function

Function CurY() As Byte
CurY = CurYY - 2000
End Function

Function CurYY%()
CurYY = Year(Now)
End Function

Function TimId$(A As Date)
TimId = Format(A, "YYYYMMDD_HHMMSS")
End Function
Function IsHMS(A$) As Boolean
If Len(A) <> 6 Then Exit Function
IsHMS = IsDate(Format(A, "00:00:00"))
End Function

Function IsDashYYMD(A$) As Boolean
Select Case True
Case Len(A) <> 10, Mid(A, 5, 1) <> "-", Mid(A, 8, 1) <> "-": Exit Function
End Select
IsDashYYMD = IsDate(A)
End Function

Function TimStr$(A As Date)
TimStr = Format(A, "YYYY-MM-DD HHMMSS")
End Function

Function NowId$()
NowId = TimId(Now)
End Function

Function IsTimStr(S) As Boolean
Select Case True
Case Len(S) <> 19, _
    Mid(S, 11, 1) <> " ", _
    Not IsDashYYMD(Left(S, 10)), _
    Not IsHMS(Right(S, 6))
    Exit Function
End Select
IsTimStr = True
End Function

Function NowStr$()
NowStr = TimStr(Now)
End Function
Function TimSi22(T As Date, Si&)
TimSi22 = Tim12(T) & AliR(Si, 10)
End Function
Function Now12$()
Now12 = Tim12(Now)
End Function
Function Tim12$(A As Date)
Tim12 = Format(A, "YYMMDDHHMMSS")
End Function

Function CvDbl(S, Optional Fun$)
Const CSub$ = CMod & "CvDbl"
'Ret : a dbl of @S if can be converted, otherwise empty and debug.print S$
On Error GoTo X
CvDbl = CDbl(S)
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dbl, emp is ret"
End Function
Sub Stamp(S)
If Not NoStamp Then Debug.Print NowStr; " "; S
End Sub
Function IsY(Y As Byte) As Boolean
IsY = True
End Function
Function IsM(M As Byte) As Boolean
IsM = True
End Function


Sub ChkIsM(M As Byte)
Raise "M should be between 1 and 12, but now[" & M & "]"
End Sub

Private Sub MonthNy__Tst()
Dim J%, A$(): A = MonthNy(19, 12)
For J = 0 To 11
    Debug.Print A(J)
Next
End Sub
Function MonthNy(Y As Byte, M As Byte, Optional NMth% = 12) As String()
Dim J%, D As Date
D = DateSerial(Y, M, 1)
For J = 0 To NMth
    PushS MonthNy, UCase(Format(D, "MMM YYYY"))
    D = FstDteNxtMth(D)
Next
End Function

Function FstDteNxtMth(D As Date) As Date
FstDteNxtMth = DateTime.DateAdd("M", 1, FstDte(D))
End Function

Function IsLasDte(D As Date) As Boolean
IsLasDte = PrvDte(FstDteNxtMth(D)) = D
End Function

Function RemDays(D As Date) As Byte
RemDays = NDay(D) - Day(D)
End Function

Function NDay(D As Date) As Byte
NDay = Day(PrvDte(FstDteNxtMth(D)))
End Function

Function RemDaysFactor!(D As Date)
Dim R As Byte: R = RemDays(D): If R = 0 Then Exit Function
RemDaysFactor = R / NDay(D)
End Function

Function NxtMth(D As Date) As Date: NxtMth = DateAdd("M", 1, D): End Function
Function PrvMth(D As Date) As Date: PrvMth = DateAdd("M", -1, D): End Function


Private Sub MthAy__Tst()
Dim MNxt() As Date: MNxt = MthAy(19, 12)
Dim MPrv() As Date: MPrv = MthAy(19, 12, -1)
Stop
End Sub
Function MthAy(Y As Byte, M As Byte, Optional Direction% = 1, Optional NMth% = 15)
Dim O() As Date: ReDim O(0 To NMth - 1)
O(0) = DateSerial(Y, M, 1)
Dim J%: For J = 1 To NMth - 1
    Select Case Direction
    Case 1: O(J) = NxtMth(O(J - 1))
    Case -1: O(J) = PrvMth(O(J - 1))
    Case Else: Stop
    End Select
Next
MthAy = O
End Function

Function IsDteAt(S, At%) As Boolean: IsDteAt = IsDate(Mid(S, At, 10)): End Function
