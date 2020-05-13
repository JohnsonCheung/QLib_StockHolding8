Attribute VB_Name = "MxDtaDaFmtDy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaFmtDy."

Sub FmtDyAsLn__Tst()
Dim A$()
A = FmtDyAsLn(SampDy3)
Dmp A
End Sub
Function FmtDyAsLn(Dy(), Optional MaxColWdt% = 100) As String()
FmtDyAsLn = FmtStrColy(StrColyzDy(StrfyDy(Dy, MaxColWdt)))
End Function
Function FmtDy(Dy(), Optional MaxColWdt% = 100, Optional Fmt As eTblFmt) As String()
FmtDy = FmtStrDy(StrfyDy(Dy, MaxColWdt), MaxColWdt, Fmt)
End Function

Private Function StrfyDy(Dy(), W%) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI StrfyDy, StrfyDr(Dr, W)
Next
End Function
Private Function StrfyDr(Dr, W%) As String()
Dim I: For Each I In Itr(Dr)
    PushI StrfyDr, StrfyVal(I, W)
Next
End Function

Function StrColyzDy(Dy()) As StrColy
Dim V()
Dim NCol%: NCol = NColzDy(Dy)
Dim J%: For J = 0 To NCol - 1
    PushI V, StrColzDy(Dy, J)
Next
StrColyzDy = StrColy(V)
End Function


Function FmtStrDy(StrDy(), Optional MaxColWdt% = 100, Optional Fmt As eTblFmt) As String()
If Si(StrDy) = 0 Then Exit Function
If W2IsLinesDy(StrDy) Then
    FmtStrDy = W2FmtLinesDy(StrDy)
Else
    FmtStrDy = W2FmtLnDy(StrDy)
End If
End Function
Private Function W2IsLinesDy(StrDy()) As Boolean
Dim Dr: For Each Dr In Itr(StrDy)
    Dim V: For Each V In Itr(Dr)
        If Not IsStr(V) Then Thw CSub, "Some ele in given StrDy is not str", "TyOf-the-ele The-ele", TypeName(V), V
        If HasLf(V) Then W2IsLinesDy = True
    Next
Next
W2IsLinesDy = False
End Function
Private Function W2FmtLinesDy(LinesDy()) As String()
'Ret : sam ele as @Dy.  One ele of @Ret may be a Ln or lines depending on any cell of Dr of @Dy has lines.
Dim W%():       W = WdtyzDy(LinesDy)
Dim O$()
Dim Sep$: Sep = Sepln(W)
PushI O, Sep '<===
Dim I: For Each I In Itr(LinesDy)
    Dim LinesDr$(): LinesDr = I
    PushIAy W2FmtLinesDy, FmtLinesDr(LinesDr, W)
Next
PushI O, Sep '<===
W2FmtLinesDy = O
End Function
Private Function W2FmtLnDy(LnDy()) As String() 'Fmt here means: Add SepUL to reach column and Ali each col.
W2FmtLnDy = AddAliStrColAv(W2AddColHdrF(StrColyzDy(LnDy)))
End Function
Private Function W2AddColHdrF(LnColy As StrColy) As Variant() ' Add :ColHdrF is Hdr-UL and Footer-UL of a column
':LnColAv: :SyAv ! #LnCol-Ay# each ele is a :LnCol
Dim LnCol: For Each LnCol In LnColy.Coly
    PushI W2AddColHdrF, W2AddColHdrFzLnCol(CvSy(LnCol))
Next
End Function
Private Function W2AddColHdrFzLnCol(LnCol$()) As String()
Dim W%: W = WdtzAy(LnCol)
Dim UL$: UL = String(W, "-")
PushI W2AddColHdrFzLnCol, UL
PushIAy W2AddColHdrFzLnCol, LnCol
PushI W2AddColHdrFzLnCol, UL
End Function

Private Function FmtLinesDr(LinesDr$(), W%()) As String()
Dim Sq(): Sq = SqzLinesDr(LinesDr)
Dim Sq1(): Sq1 = AliSqW(Sq, W)
Dim Dr, Ln$, IR%: For IR = 1 To UB1(Sq)
    Dr = DrzSq(Sq1, IR)
    Ln = TblFmtLn(Dr)
    PushI FmtLinesDr, Ln
Next
End Function

Private Function SqzLinesDr(LinesDr$()) As Variant()
Dim NRow%: NRow = MaxLnCnt(LinesDr): If NRow = 0 Then Exit Function
Dim NCol%: NCol = Si(LinesDr)
Dim Coly(): Coly = ColyzLinesDr(LinesDr)
Dim O(): ReDim O(1 To NRow, 1 To NCol)
Dim Col$(), ICol%, S$: For ICol = 0 To NCol - 1
    Col = Coly(ICol)
    Dim IRow%: For IRow = 0 To UB(Col)
        O(IRow + 1, ICol + 1) = Col(IRow)
    Next
Next
SqzLinesDr = O
End Function

Function ColyzLinesDr(LinesDr$()) As Variant()
Dim Lines: For Each Lines In LinesDr
    PushI ColyzLinesDr, SplitCrLf(Lines)
Next
End Function

Private Function WdtyzLinesDy(LinesDy()) As Integer()
Dim Dr, LinesDr$(): For Each Dr In Itr(LinesDy)
    LinesDr = Dr
    WdtyzLinesDy = AddWdty(WdtyzLinesDy, WdtyzLinesDr(LinesDr))
Next
End Function

Private Function WdtyzLinesDr(LinesDr$()) As Integer() 'Return Wdty of each ele of @LinesDr, which is a Lines
Dim Lines: For Each Lines In Itr(LinesDr)
    PushI WdtyzLinesDr, WdtzLines(Lines)
Next
End Function

Function AddWdty(Wdty1%(), Wdty2%()) As Integer()
Dim MinU%, MaxU%, U1%, U2%, O%()
U1 = UB(Wdty1)
U2 = UB(Wdty2)
MinU = Min(U1, U2)
MaxU = Max(U1, U2)
O = Wdty1
Dim J%: For J = 0 To MinU
    If Wdty2(J) > Wdty1(J) Then O(J) = Wdty2(J)
Next
If U2 > U1 Then
    ReDim Preserve O(MaxU)
    For J = MinU + 1 To MaxU
        O(J) = Wdty2(J)
    Next
End If
AddWdty = O
End Function

Function ShwZerDrs(A As Drs, ShwZer As Boolean) As Drs
If ShwZer Then ShwZerDrs = A: Exit Function
Dim ODy()
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim J%: For J = 0 To UB(Dr)
            If Dr(J) = "0" Then Dr(J) = ""
        Next
        PushI ODy, Dr
    Next
ShwZerDrs = Drs(A.Fny, ODy)
End Function

Function Sepln$(W%(), Optional NoColSep As Boolean)
Dim Q$(): Q = ColSepAy(W, Not NoColSep)
Sepln = FmtLn(Q, W, SepDr(W))
End Function

Private Function SepDr(W%()) As String()
Dim I: For Each I In Itr(W)
    Push SepDr, Dup("-", I)
Next
End Function

Function TblFmtLn$(Dr)
TblFmtLn = "| " & Jn(Dr, " | ") & " |"
End Function
