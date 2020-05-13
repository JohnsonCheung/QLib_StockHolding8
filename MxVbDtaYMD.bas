Attribute VB_Name = "MxVbDtaYMD"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "MxVbDtaYMD."
Public Type Ymd
    Y As Byte
    M As Byte
    D As Byte
End Type
Type YM
    Y As Byte
    M As Byte
End Type

Function YMzYmd(A As Ymd) As YM
YMzYmd = YM(A.Y, A.M)
End Function
Function YMzYYMM(YYMM%) As YM
YMzYYMM = YM(YYMM \ 100, YYMM Mod 100)
End Function
Function YM(Y As Byte, M As Byte) As YM
YM.Y = Y
YM.M = M
End Function
Function YM_IsEmpty(A As YM) As Boolean
YM_IsEmpty = A.Y = 0
End Function
Function YYmdStr$(A As Ymd)
With A
YYmdStr = .Y + 2000 & "-" & Format(.M, "00") & "-" & Format(.D, "00")
End With
End Function
Function YmdHStr$(A As Ymd)
With A
YmdHStr = .Y & "-" & Format(.M, "00") & "-" & Format(.D, "00")
End With
End Function
Function YmdUB&(A() As Ymd)
YmdUB = YmdSi(A) - 1
End Function
Function YmdSi&(A() As Ymd)
On Error Resume Next
YmdSi = UBound(A) + 1
End Function

Sub PushYmd(O() As Ymd, M As Ymd)
If IsEmpYmd(M) Then Exit Sub
Dim N&: N = YmdSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Function IsEmpYmd(A As Ymd) As Boolean
With A
Select Case True
Case .Y = 0, .M = 0, .D = 0: IsEmpYmd = True
End Select
End With
End Function
Function Ymd(Y As Byte, M As Byte, D As Byte) As Ymd
With Ymd
.Y = Y
.M = M
.D = D
End With
End Function
Function YmdzYYMMDD(YYMMDD&) As Ymd
With YmdzYYMMDD
    .Y = YYMMDD \ 10000
    .M = (YYMMDD - .Y * 10000&) \ 100
    .D = YYMMDD Mod 100
End With
End Function

Function YYMMDD&(Y As Byte, M As Byte, D As Byte)
YYMMDD = CLng(Y) * 10000 + CLng(M) * 100 + D
End Function

Function YmdzDte(D As Date) As Ymd
With YmdzDte
    .Y = Year(D) - 2000
    .M = Month(D)
    .D = Day(D)
End With
End Function

Function DtezYmd(A As Ymd) As Date
With A
DtezYmd = DateSerial(.Y, .M, .D)
End With
End Function

Function YymStr$(Y As Byte, M As Byte)
YymStr = 2000 + Y & "-" & Format(M, "00")
End Function
Function YymNumStr$(Y As Byte, M As Byte)
YymNumStr = 2000 + Y & Format(M, "00")
End Function

Function YYMStrzYM$(A As YM)
YYMStrzYM = YymStr(A.Y, A.M)
End Function
