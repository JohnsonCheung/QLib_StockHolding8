Attribute VB_Name = "MxDtaDaCsv"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaCsv."

Function DrzCsl(Csl) As Variant()
If Not HasDblQ(Csl) Then
    DrzCsl = Split(Csl, ",")
    Exit Function
End If
Dim L$: L = Trim(Csl)
Dim J%
While L <> ""
    LoopTooMuch CSub, J
    PushI DrzCsl, ShfCsvTerm(L)
Wend
End Function

Function SyzCsl(Csl) As String()
If Not HasDblQ(Csl) Then
    SyzCsl = SplitComma(Csl)
    Exit Function
End If
Dim L$: L = Trim(Csl)
Dim J%
While L <> ""
    LoopTooMuch CSub, J
    PushI SyzCsl, ShfCsvTerm(L)
Wend
End Function

Function ShfCsvTerm$(OLin$)
Dim NotDblQ As Boolean
Dim DblQCommaPos As Boolean
Dim LasIsDblQ As Boolean
    NotDblQ = FstChr(OLin) <> vbDblQ
    If Not NotDblQ Then DblQCommaPos = InStr(2, OLin, vbDblQ & vbComma)
    If DblQCommaPos = 0 Then LasIsDblQ = LasChr(OLin) = vbDblQ

Select Case True
Case NotDblQ
    ShfCsvTerm = BefCommaOrAll(OLin)
    OLin = AftComma(OLin)
Case DblQCommaPos > 0
    ShfCsvTerm = Replace(Mid(OLin, 2, DblQCommaPos - 1), vb2DblQ, vbDblQ)
    OLin = Mid(OLin, DblQCommaPos + 1)
Case LasIsDblQ
    ShfCsvTerm = Replace(Mid(OLin, 2, DblQCommaPos - 1), vb2DblQ, vbDblQ)
    OLin = ""
Case Else
    Thw CSub, "CsvEr: OLin has FstChr is DblQ, No DblQComm, Las<>DblQ, it should be Closing-vbDblQ", "OLin", OLin
End Select
End Function

Function QuoCsv$(V)
Select Case True
Case IsStr(V): QuoCsv = QuoDbl(Replace(V, vbDblQ, vb2DblQ))
Case IsDte(V): QuoCsv = "#" & Format(V, "YYYY-MM-DD HH:MM:SS") & "#"
Case IsEmpty(V):
Case Else: QuoCsv = V
End Select
End Function

Function Csl$(Dr)
':Csl: :Ln ! #Comma-Separated-Line#
Dim U%: U = UB(Dr)
If U = -1 Then Exit Function
Dim J&
Dim O$(): ReDim O(U)
Dim V: For Each V In Dr
    O(J) = QuoCsv(V)
    J = J + 1
Next
Csl = JnComma(O)
End Function
