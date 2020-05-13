Attribute VB_Name = "MxXlsRgOffset"
Option Explicit
Option Compare Text
Const CNs$ = "Xls.Rg"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRgOffset."

Function C2zRg%(Rg As Range):                                C2zRg = Rg.Column + Rg.Columns.Count - 1:                 End Function
Function EntRgC(A As Range, C) As Range:                Set EntRgC = RgC(A, C).EntireColumn:                           End Function
Function EntRgCC(A As Range, C1, C2) As Range:         Set EntRgCC = RgCC(A, C1, C2).EntireColumn:                     End Function
Function EntRgR(A As Range, Optional R = 1) As Range:             Set EntRgR = RgR(A, R).EntireRow:                              End Function
Function EntRgRR(A As Range, R1, R2) As Range:         Set EntRgRR = RgRR(A, R1, R2).EntireRow:                        End Function
Function FstColRg(A As Range) As Range:               Set FstColRg = RgC(A, 1):                                        End Function
Function FstDtaCnozRg%(Rg As Range):                  FstDtaCnozRg = FstDtaCno(WszRg(Rg), RCCzRg(Rg)):                 End Function
Function FstRowRg(A As Range) As Range:               Set FstRowRg = RgR(A, 1):                                        End Function
Function RCCzRg(Rg As Range) As RCC:                        RCCzRg = RCC(Rg.Row, Rg.Column, C2zRg(Rg)):                End Function
Function RgC(A As Range, Optional C = 1) As Range:         Set RgC = RgCC(A, C, C):                                    End Function
Function RgCC(A As Range, C1, C2) As Range:               Set RgCC = RgRCRC(A, 1, C1, NRoZZRg(A), C2):                 End Function
Function RgCRR(A As Range, C, R1, R2) As Range:          Set RgCRR = RgRCRC(A, R1, C, R2, C):                          End Function
Function RgMoreBelow(A As Range, Optional N% = 1): Set RgMoreBelow = RgRR(A, 1, NRoZZRg(A) + N):                       End Function
Function RgR(A As Range, Optional R = 1) As Range:         Set RgR = RgRR(A, R, R):                                    End Function
Function RgRC(A As Range, R, C) As Range:                 Set RgRC = A.Cells(R, C):                                    End Function
Function RgRCC(A As Range, R, C1, C2) As Range:          Set RgRCC = RgRCRC(A, R, C1, R, C2):                          End Function
Function RgRCRC(A As Range, R1, C1, R2, C2) As Range:   Set RgRCRC = WszRg(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2)): End Function
Function RgRR(A As Range, R1, R2) As Range:               Set RgRR = RgRCRC(A, R1, 1, R2, NColzRg(A)):                 End Function

Function RgMoreTop(A As Range, Optional N = 1)
Dim O As Range
Set O = RgRR(A, 1 - N, NRoZZRg(A))
Set RgMoreTop = O
End Function

Function FstDtaCno%(Ws As Worksheet, RCC As RCC)
With RCC
Dim Cno%: For Cno = .C1 To .C2
    If Not IsEmpty(WsRC(Ws, .R, Cno).Value) Then FstDtaCno = Cno: Exit Function
Next
End With
End Function
