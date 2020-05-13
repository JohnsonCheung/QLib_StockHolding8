Attribute VB_Name = "gzDta"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzDta."
Type CoYM
    Co As Byte
    M As Byte
    Y As Byte
End Type
Type StmYM
    Stm As String
    Y As Byte
    M As Byte
End Type
Type CoYmd
    Co As Byte
    Ymd As Ymd
End Type
Type CoStmYM
    Co As Byte
    Stm As String
    Y As Byte
    M As Byte
End Type
Function CoYmd(Co As Byte, A As Ymd) As CoYmd
ChkIsCo Co
CoYmd.Co = Co
CoYmd.Ymd = A
End Function

Function CoYM(Co As Byte, Y As Byte, M As Byte) As CoYM
ChkIsCo Co
CoYM.Co = Co
CoYM.Y = Y
CoYM.M = M
End Function

Function CoYMzCoYmd(A As CoYmd) As CoYM
With CoYMzCoYmd
    .Co = A.Co
    .Y = A.Ymd.Y
    .M = A.Ymd.M
End With
End Function
Function CoYmdStr$(A As CoYmd)
CoYmdStr = YYmdStr(A.Ymd) & "(" & A.Co & "00" & ")"
End Function

Function StmYmStr$(A As StmYM)
With A
StmYmStr = "Stm[" & .Stm & "] Y[" & .Y & "] M[" & .M & "]"
End With
End Function

Function StmYmNumStr$(A As StmYM)
With A
StmYmNumStr = YymNumStr(A.Y, A.M) & Stm2zStm(A.Stm)
End With
End Function

Function StmYM(Stm$, Y As Byte, M As Byte) As StmYM
ChkIsStm Stm
Select Case True
Case Y > 99:         PmEr "StmYM", "Y", Y, "0-99"
End Select
With StmYM
    .Stm = Stm
    .Y = Y
    .M = M
End With
End Function
Function StmYM1(Stm$, A As YM) As StmYM
With StmYM1
    .Stm = Stm
    .Y = A.Y
    .M = A.M
End With
End Function

Function StmYMSi%(A() As StmYM)
On Error Resume Next
StmYMSi = UBound(A) + 1
End Function
Function IsEmpStmYM(A As StmYM) As Boolean
With A
Select Case True
Case Not IsStm(.Stm), Not IsY(.Y), Not IsM(.M): IsEmpStmYM = True
End Select
End With
End Function

Sub PushNonEmpStmYM(O() As StmYM, M As StmYM)
If Not IsEmpStmYM(M) Then PushStmYM O, M
End Sub
Sub PushStmYM(O() As StmYM, M As StmYM)
Dim N%: N = StmYMSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub RaiseStm(Stm$): Raise "Stm should be U or M, but now[" & Stm & "]": End Sub
Sub RaiseCo(Co As Byte): Raise "Co should be 86 or 87, but now[" & Co & "]": End Sub
Sub ChkIsCo(Co As Byte)
If Co <> 86 And Co <> 87 Then RaiseCo Co
End Sub
Sub ChkIsStm(Stm$)
If Not IsStm(Stm) Then RaiseStm Stm
End Sub
Function IsStm(Stm$) As Boolean
Select Case Stm
Case "U", "M": IsStm = True
End Select
End Function

Function CoStr$(Co As Byte)
CoStr = Co & "00"
End Function
Function CoNm$(Co As Byte)
Select Case Co
Case 86: CoNm = "HK"
Case 87: CoNm = "Macau"
Case Else: Co = "?"
End Select
End Function
Function StmzStm2$(Stm2$)
Dim O$
Select Case Stm2
Case "MH": O = "M"
Case "UD": O = "U"
End Select
StmzStm2 = O
End Function

Function Stm2zStm$(Stm$)
Select Case Stm
Case "U": Stm2zStm = "UD"
Case "M": Stm2zStm = "MH"
Case Else: RaiseStm Stm
End Select
End Function
Function StreamzStm$(Stm$)
Select Case Stm
Case "U": StreamzStm = "Diageo"
Case "M": StreamzStm = "MH"
Case Else: RaiseStm Stm
End Select
End Function
Function StmzTopz$(CdTopaz)
If IsNull(CdTopaz) Then StmzTopz = "?": Exit Function
StmzTopz = IIf(Left(CdTopaz, 3) = "UDV", "U", "M")
End Function

Function CoStmYM(Co As Byte, Stm$, Y As Byte, M As Byte) As CoStmYM
With CoStmYM
    .Co = Co
    .Stm = Stm
    .M = M
    .Y = Y
End With
End Function

Function CoStmYM_DtaStr$(A As CoStmYM)
With A
CoStmYM_DtaStr = "Co[" & .Co & "] Stm[" & .Stm & "] Y[" & .Y & "] M[" & .M & "]"
End With
End Function
