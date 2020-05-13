Attribute VB_Name = "MxXlsLof"
Option Explicit
Option Compare Text
Const CNs$ = "Lof"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLof."
Public Const LofAliss$ = ""
Public Const LofBdrss$ = ""

Private Type AdjLyEr
    AdjLy() As String
    Er() As String
End Type
Private Type Vdt
    Er() As String
    Lon As String
    Fny() As String
    Ali() As String
    Wdt() As String
    Bdr() As String
    Lvl() As String
    Cor() As String
    Tot() As String
    Fmt() As String
    Tit() As String
    Fml() As String
    Lbl() As String
    Sum() As String
End Type
Enum eLofAli: EiLofAliL: EiAliC: EiAliR: End Enum
Enum eLofBdr: EiLofBdrL: EiBdrC: EiBdrRt: End Enum
Enum eLofTot: EiLofTotSum: EiLofAvg: EiLofCntt: End Enum
Type LofAli: Fny() As String: Ali As eLofAli: End Type
Type LofWdt: Fny() As String: Wdt As Integer:  End Type
Type LofBdr: Fny() As String: Bdr As eLofBdr: End Type
Type LofLvl: Fny() As String: Lvl As Byte:     End Type
Type LofCor: Fny() As String: Cor As Long:     End Type
Type LofTot: Fny() As String: Tot As eLofTot: End Type
Type LofFmt: Fny() As String: Fmt As String:   End Type
Type LofTit: Fld   As String: Tit As String:   End Type
Type LofFml: Fld   As String: Fml As String:   End Type
Type LofLbl: Fld   As String: Lbl As String:   End Type
Type LofSum: SumFld As String: FmFld As String: ToFld As String: End Type
Type Lof
    Lon As String
    Fny() As String
    Ali() As LofAli
    Bdr() As LofBdr
    Cor() As LofCor
    Fml() As LofFml
    Fmt() As LofFmt
    Lbl() As LofLbl
    Lvl() As LofLvl
    Sum() As LofSum
    Tit() As LofTit
    Tot() As LofTot
    Wdt() As LofWdt
End Type

':FunPfx-Do: :FunPfx ! #FunPfx-Drs-Of#
Private Sub Lof__Tst()
Dim A As Lof: A = Lof(SampLofLy)
Stop
End Sub

Function Lof(LofLy$()) As Lof
Dim A As Vdt: A = VdtLof(LofLy)
ChkEr A.Er, "Lof"
With Lof
    .Fny = A.Fny
    .Lon = A.Lon
    .Ali = BrkAli(A.Ali)
'    .Wdt = BrkWdt(A.Wdt)
'    .Bdr = BrkBdr(A.Bdr)
'    .Lvl = BrkLvl(A.Lvl)
'    .Cor = BrkCor(A.Cor)
'    .Tot = Brktot(A.Tot)
'    .Fmt = BrkFmt(A.Fmt)
'    .Tit = BrkTit(A.Tit)
'    .Fml = BrkFml(A.Fml)
'    .Lbl = BrkLbl(A.Lbl)
'    .Sum = BrkSum(A.Sum)
End With
End Function

Private Function VdtLof(LofLy$()) As Vdt
Dim O As Vdt, L$()
L = LofLy
With ShfAdjFny(L)
    
    
End With
VdtLof = O
End Function

Private Function ShfAdjFny(OLy$()) As AdjLyEr
'@LttdDrs                 ! Select T1="Lo" and FstTerm(*Dta)='Fld' and Set *LoFny = SyzSS(Rst(*Dta))
'Ret     :Drs-L-LoFny
End Function

Private Function Do_L_Fm_To_Sum(LttdDrs As Drs) As Drs
'@LttdDrs ! *T1.T2 = 'Sum Bet' *Dta = *Fm *To *Sum
'Ret  :Drs-L-Fm-To-Sum
Dim Dy()
    Dim Dr: For Each Dr In Itr(DwCC2EqExl(LttdDrs, "T1 T2", "Sum", "Bet").Dy)
        Dim L&: L = Dr(0)
        Dim Dta$: Dta = Dr(1)
        Dim FldFm$
        Dim FldTo$
        Dim FldSum$: AsgTTRst Dta, FldFm, FldTo, FldSum
        PushI Dy, Array(L, FldFm, FldTo, FldSum)
    Next
Do_L_Fm_To_Sum = DrszFF("L Fm To Sum", Dy)
End Function
Private Function LofAliUB&(A() As LofAli): On Error Resume Next: LofAliUB = UBound(A) + 1: End Function
Private Function LofTitUB&(A() As LofTit): On Error Resume Next: LofTitUB = UBound(A) + 1: End Function
Private Function LofAliSi&(A() As LofAli): LofAliSi = LofAliUB(A) - 1: End Function
Private Function LofTitSi&(A() As LofTit): LofTitSi = LofTitUB(A) - 1: End Function
Private Sub PushLofAli(O() As LofAli, M As LofAli): Dim N&: N = LofAliSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Sub PushLofTit(O() As LofTit, M As LofTit): Dim N&: N = LofTitSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function BrkAli(Ali$()) As LofAli()
Dim L: For Each L In Itr(Ali)
    PushLofAli BrkAli, BrkOneAli(L)
Next
End Function
Private Function BrkOneAli(Ali) As LofAli
End Function

Function LofTitAy(LofTitLy$()) As LofTit()
Dim L: For Each L In Itr(LofTitLy)
    PushLofTit LofTitAy, LofTitzLin(L)
Next
End Function

Function LofTitzLin(LofTitLin) As LofTit
Dim O As LofTit
With BrkSpc(LofTitLin)
    O.Fld = .S1
    O.Tit = .S2
End With
End Function
