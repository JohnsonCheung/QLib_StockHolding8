Attribute VB_Name = "MxDtaAgr"
Option Explicit
Option Compare Text
Const CNs$ = "Dta.Op"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaAgr."

Function AgrDrs(D As Drs, Gpcc$, ArgColn$) As Drs
Dim Dy()
    Dim A As Drs: A = Gp(D, Gpcc, ArgColn)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Col(): Col = Pop(Dr)
        Dim Sum#: Sum = AySum(Col)
        Dim N&:   N = Si(Col)
        Dim Avg: If N <> 0 Then Avg = Sum / N
        PushI Dr, N
        PushI Dr, Avg
        PushI Dr, AySum(Col)
        PushI Dr, MinEle(Col)
        PushI Dr, MaxEle(Col)
        PushI Dy, Dr
    Next
Dim NewFny$(): NewFny = SyzSS(RplQ("?Cnt ?Avg ?Sum ?Min ?Max", ArgColn))
Dim Fny$(): Fny = AddSy(D.Fny, NewFny)
AgrDrs = Drs(Fny, Dy)
End Function

Private Sub AgrDrsWdt__Tst()
BrwDrs AgrDrsWdt(PFunDrsP, "Mdn Ty", "Mthn")
End Sub

Function AgrDrsWdt(D As Drs, Gpcc$, C$) As Drs
Dim A As Drs: A = Gp(D, Gpcc, C)
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dim Col(): Col = Pop(Dr)
    PushI Dr, WdtzAy(Col)
    PushI Dy, Dr
Next
Dim Fny$(): Fny(UB(Fny)) = "W" & C
AgrDrsWdt = Drs(Fny, Dy)
End Function

Function AgrDrsMin(D As Drs, Gpcc$, MinC$) As Drs
Dim Dy()
    Dim A As Drs: A = Gp(D, Gpcc, MinC)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Col(): Col = Pop(Dr)
        PushI Dr, MinEle(Col)
        PushI Dy, Dr
    Next
AgrDrsMin = Drs(D.Fny, Dy)
End Function

Function AgrDrsMax(D As Drs, Gpcc$, MaxC$) As Drs
Dim Dy()
    Dim A As Drs: A = Gp(D, Gpcc, MaxC)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Col(): Col = Pop(Dr)
        PushI Dr, MaxEle(Col)
        PushI Dy, Dr
    Next
AgrDrsMax = Drs(D.Fny, Dy)
End Function
