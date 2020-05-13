Attribute VB_Name = "MxDtaDaSq"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CNs$ = "Dta.Sq"
Const CMod$ = CLib & "MxDtaDaSq."
Function SqzDy(Dy(), Optional SKipNRow& = 1) As Variant()
Dim O(), NR&, NC&
NR = Si(Dy)
NC = NColzDy(Dy)
ReDim O(1 To NR, 1 To NC)
Dim R&: For R = 1 To NR
    Dim Dr: Dr = Dy(R - 1)
    SetSqr O, Dr, R
Next
SqzDy = O
End Function

Sub SetSqr(OSq(), Dr, Optional R = 1, Optional NoTxtSngQ As Boolean)
Dim J&
If NoTxtSngQ Then
    For J = 0 To UB(Dr)
        If IsStr(Dr(J)) Then
            OSq(R, J + 1) = QuoSng(CStr(Dr(J)))
        Else
            OSq(R, J + 1) = Dr(J)
        End If
    Next
Else
    For J = 0 To UB(Dr)
        OSq(R, J + 1) = Dr(J)
    Next
End If
End Sub

Sub PushSq(OSq(), Sq())
Const CSub$ = CMod & "PushSq"
Dim NR&: NR = UBound(OSq, 1) + UBound(Sq, 1)
Dim NC&: NC = UBound(OSq, 2)
Dim NC2&: NC2 = UBound(Sq, 2)
If NC <> NC2 Then Thw CSub, "NC of { OSq, Sq } are dif", "OSq-NC Sq-NC", NC, NC2
ReDim Preserve OSq(1 To NR, 1 To NC)
Dim R&, C&
For R = 1 To NC2
    For C = 1 To NC
        OSq(R + NR, C) = Sq(R, C)
    Next
Next
End Sub

Function Sq(R&, C&) As Variant()
'Ret : a Sq(1 to @R, 1 to @C)
Dim O()
ReDim O(1 To R, 1 To C)
Sq = O
End Function

Function AddSngQuozSq(Sq())
Dim NC%, C%, R&, O
O = Sq
NC = UBound(Sq, 2)
For R = 1 To UBound(Sq, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
AddSngQuozSq = O
End Function
Function JnSq(Sq(), SepChr$) As String()
Dim NC&: NC = UBound(Sq, 2)
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI JnSq, JnSqr(Sq, R, SepChr)
Next
End Function

Function JnSqr$(Sq(), R&, SepChr$)
JnSqr = Join(SyzSqr(Sq, R), SepChr)
End Function

Function SyzSqr(Sq(), R&) As String()
Dim J&
For J = 1 To UBound(Sq, 2)
    PushI SyzSqr, Sq(R, J)
Next
End Function

Function FmtSq(Sq(), Optional SepChr$ = " ") As String()
FmtSq = JnSq(AliSq(Sq), SepChr)
End Function

Sub BrwSq(Sq())
Brw FmtSq(Sq)
End Sub

Function ColzSq(Sq(), Optional C = 1) As Variant()
ColzSq = IntozSqc(EmpAv, Sq, C)
End Function

Function DrzSq(Sq(), Optional R = 1) As Variant()
DrzSq = IntozSqr(EmpAv, Sq, R)
End Function

Function IntozSqc(Into, Sq(), C)
Dim NR&: NR = UBound(Sq, 1)
Dim O:    O = ResiN(Into, NR)
Dim R&: For R = 1 To NR
    O(R - 1) = Sq(R, C)
Next
IntozSqc = O
End Function

Function F_Into_SelSq_ByR_AndColnoy(Into, Sq(), R, Colnoy%())
Dim NCol&:    NCol = UBound(Colnoy)
Dim O: O = Into: ReDim O(NCol - 1)
Dim C%: For C = 0 To NCol - 1
    O(C) = Sq(R, Colnoy(C))
Next
F_Into_SelSq_ByR_AndColnoy = O
End Function

Function IntozSqr(Into, Sq(), R)
Dim NCol&:    NCol = UBound(Sq, 2)
Dim O: O = Into: ReDim O(NCol - 1)
Dim C%: For C = 1 To NCol
    O(C - 1) = Sq(R, C)
Next
IntozSqr = O
End Function

Function IntozSqrColnoy(Into, Sq(), R, Colnoy)
Dim UCol%:    UCol = UBound(Colnoy)
Dim O: O = Into: ReDim O(UCol)
Dim C%: For C = 1 To UCol + 1
    O(C - 1) = Sq(R, C)
Next
IntozSqrColnoy = O
End Function

Function SyzSq(Sq(), Optional C& = 1) As String()
SyzSq = IntozSqc(EmpSy, Sq(), C)
End Function

Function DrzSqr(Sq(), Optional R = 1) As Variant()
DrzSqr = IntozSqr(EmpAv, Sq, R)
End Function

Function DrzSqrColnoy(Sq(), R, Colnoy%()) As Variant()
DrzSqrColnoy = IntozSqrColnoy(EmpAv, Sq, R, Colnoy)
End Function

Function F_Dr_SelSq_ByR_AndColnoy(Sq(), R, Colnoy%()) As Variant()
F_Dr_SelSq_ByR_AndColnoy = F_Into_SelSq_ByR_AndColnoy(EmpAv, Sq, R, Colnoy)
End Function

Function InsSqr(Sq(), Dr(), Optional Row& = 1)
Dim O(), C&, R&, NC&, NR&
NC = NColzSq(Sq)
NR = NRowOfSq(Sq)
ReDim O(1 To NR + 1, 1 To NC)
For R = 1 To Row - 1
    For C = 1 To NC
        O(R, C) = Sq(R, C)
    Next
Next
For C = 1 To NC
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To NC
        O(R + 1, C) = Sq(R, C)
    Next
Next
InsSqr = O
End Function

Function IsEqSq(A, B) As Boolean
Dim NR&, NC&
NR = UBound(A, 1)
NC = UBound(A, 2)
If NR <> UBound(B, 1) Then Exit Function
If NC <> UBound(B, 2) Then Exit Function
Dim R&, C&
For R = 1 To NR
    For C = 1 To NC
        If A(R, C) <> B(R, C) Then
            Exit Function
        End If
    Next
Next
IsEqSq = True
End Function

Function TermLnAyzSq(Sq()) As String()
Dim R&
For R = 1 To UBound(Sq(), 1)
    Push TermLnAyzSq, Termln(DrzSqr(Sq, R))
Next
End Function

Function NColzSq&(Sq())
On Error Resume Next
NColzSq = UBound(Sq, 2)
End Function
Function NwLoSqAt(Sq(), At As Range) As ListObject
Set NwLoSqAt = NwLo(RgzSq(Sq(), At))
End Function

Function NwLoSq(Sq(), Optional Wsn$ = "Data") As ListObject
Set NwLoSq = NwLoSqAt(Sq(), NwA1(Wsn))
End Function

Function WszSq(Sq(), Optional Wsn$) As Worksheet
Set WszSq = NwLo(RgzSq(Sq(), NwA1(Wsn)))
End Function

Function NRowOfSq&(Sq())
On Error Resume Next
NRowOfSq = UBound(Sq, 1)
End Function


Function Transpose(Sq()) As Variant()
Dim NR&, NC&
NR = NRowOfSq(Sq): If NR = 0 Then Exit Function
NC = NColzSq(Sq): If NC = 0 Then Exit Function
Dim O(), J&, I&
ReDim O(1 To NC, 1 To NR)
For J = 1 To NR
    For I = 1 To NC
        O(I, J) = Sq(J, I)
    Next
Next
Transpose = O
End Function


Function SampSq() As Variant()
Const NR% = 10
Const NC% = 10
Dim O(), R%, C%
ReDim O(1 To NR, 1 To NC)
SampSq = O
For R = 1 To NR
    For C = 1 To NC
        O(R, C) = R * 1000 + C
    Next
Next
SampSq = O
End Function

Function CvDte(S, Optional Fun$)
Const CSub$ = CMod & "CvDte"
'Ret : a date fm @S if can be converted, otherwise empty and debug.print @S
On Error GoTo X
Dim O As Date: O = S
If CntSubStr(S, "/") <> 2 Then GoTo X ' ! one [/]-str is cv to yyyy/mm, which is not consider as a dte.
'                                       ! so use 2-[/] to treat as a dte str.
If Year(O) < 2000 Then GoTo X         ' ! year < 2000, treat it as str or not
CvDte = O
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dte, emp is ret"
End Function
Private Sub SqStr__Tst()
Brw SqStrzDrs(MthDrsP)
End Sub
Function SqStrzDy$(Dy())
SqStrzDy = SqStr(SqzDy(Dy))
End Function
Function SqStrzDrs$(D As Drs)
SqStrzDrs = SqStrzDy(D.Dy)
End Function

Function SqStrzWs$(S As Worksheet)
SqStrzWs = SqStrzLo(FstLo(S))
End Function

Function SqStrzLo$(L As ListObject)
SqStrzLo = SqStrzRg(L.DataBodyRange)
End Function

Function CellStr$(V, Optional Fun$)
':CellStr: :S #Xls-Cell-Str# ! A str coming fm xls cell
Dim T$: T = TypeName(V)
Dim O$
Select Case T
Case "Boolean", "Long", "Integer", "Date", "Currency", "Single", "Double": CellStr = V
Case "String": If IsDblStr(V) Then CellStr = "'" & V Else CellStr = SlashCrLfTab(V)
Case Else: If Fun <> "" Then Inf Fun, "Val-of-TypeName[" & T & "] cannot cv to :CellStr"
End Select
End Function

Function SqStrzRg$(R As Range)
SqStrzRg = SqStr(SqzRg(R))
End Function

Function IsSqEmp(Sq()) As Boolean
Dim R&: For R = 1 To UBound(Sq, 1)
    Dim C%: For C = 1 To UBound(Sq, 2)
        If Not IsEmpty(Sq(R, C)) Then Exit Function
    Next
Next
IsSqEmp = True
End Function

Function DrszSq(SqWiHdr()) As Drs
Dim Fny$(): Fny = SyzSqr(SqWiHdr, 1)
Dim Dy()
    Dim R&: For R = 2 To UBound(SqWiHdr, 1)
        PushI Dy, DrzSqr(SqWiHdr, R)
    Next
DrszSq = Drs(Fny, Dy)
End Function

Function FstDrzSq(Sq()) As Variant()
Dim O(): ReDim O(UBound(Sq, 2) - 1)
Dim J%: For J = 0 To UBound(O)
    O(J) = Sq(1, J + 1)
Next
FstDrzSq = O
End Function

Function FstColzSq(Sq()) As Variant()
Dim O(): ReDim O(UBound(Sq, 2) - 1)
Dim J&: For J = 0 To UBound(O)
    O(J) = Sq(J + 1, 1)
Next
FstColzSq = O
End Function

Property Get SampSq1() As Variant()
Dim O(), R&, C&
Const NR& = 1000
Const NC& = 100
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
For C = 1 To NC
    O(R, C) = R + C
Next
Next
SampSq1 = O
End Property
Property Get SampSqWithHdr() As Variant()
SampSqWithHdr = InsSqr(SampSq, SampDr_AToJ)
End Property
