Attribute VB_Name = "MxVbDtaSqFmt"
Option Explicit
Option Compare Text
Const CNs$ = "Dta.Sq.Op"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbDtaSqFmt."

Function AliSqW(Sq(), W%()) As Variant()
Dim O(): O = Sq
Dim C%: For C = 1 To UB2(Sq)
    Dim Wdt%: Wdt = W(C - 1)
    Dim IR&: For IR = 1 To UB1(Sq)
        O(IR, C) = Ali(O(IR, C), Wdt)
    Next
Next
AliSqW = O
End Function

Function AliSq(Sq()) As Variant()
AliSq = AliSqW(Sq, WdtyzSq(Sq))
End Function

Private Function WdtyzSq(Sq()) As Integer()
Dim C%: For C = 1 To UBound(Sq, 2)
    PushI WdtyzSq, WdtzSqc(Sq, C)
Next
End Function

Private Function WdtzSqc%(Sq(), C%)
Dim R&, O%: For R = 1 To UBound(Sq, 1)
    O = Max(O, Len(Sq(R, C)))
Next
WdtzSqc = O
End Function
