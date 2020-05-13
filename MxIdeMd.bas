Attribute VB_Name = "MxIdeMd"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMd."
Type LLn: Lno As Integer: Ln As String: End Type ' Deriving(Ay Ctor)
':DiMdDnqSrcl: :Dic ! It is from Pj. Key is Mdn and Val is MdLnes"
':MdDn: :Pjn.Mdn|Mdn

Sub ChkMdLno(Fun$, M As CodeModule, Lno&)
If Not IsBet(Lno, 1, M.CountOfLines) Then
    Thw Fun, "Lno is out of Md range", "Mdn Lno Md-Max-Lno", Mdn(M), Lno, M.CountOfLines
End If
End Sub

Function MdDic(P As VBProject, Optional MdnPatn$ = ".*") As Dictionary
Set MdDic = New Dictionary
Dim R As RegExp: Set R = Rx(MdnPatn)
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMch(C.Name, R) Then
        MdDic.Add C.Name, Srcl(C.CodeModule)
    End If
Next
End Function
Function MdDicP(Optional MdnPatn$ = ".*") As Dictionary: Set MdDicP = MdDic(CPj, MdnPatn): End Function

Function Md(MdDn) As CodeModule
Const CSub$ = CMod & "Md"
Dim A1$(): A1 = Split(MdDn, ".")
Select Case Si(A1)
Case 1: Set Md = CPj.VBComponents(A1(0)).CodeModule
Case 2: Set Md = Pj(A1(0)).VBComponents(A1(1)).CodeModule
Case Else: Thw CSub, "[MdDn] should be XXX.XXX or XXX", "MdDn", MdDn
End Select
End Function

Function MdDn$(M As CodeModule)
MdDn = PjnzM(M) & "." & Mdn(M)
End Function

Function DiMdnqSrclzP(P As VBProject) As Dictionary
Dim C As VBComponent
Set DiMdnqSrclzP = New Dictionary
For Each C In P.VBComponents
    DiMdnqSrclzP.Add C.Name, Srcl(C.CodeModule)
Next
End Function

Function DiMdnqSrclP() As Dictionary
Set DiMdnqSrclP = DiMdnqSrclzP(CPj)
End Function

Function MdFn$(M As CodeModule)
MdFn = Mdn(M) & ExtzCmpTy(CmpTyzM(M))
End Function

Function Mdn(M As CodeModule)
Mdn = M.Name
End Function

Function MdnzM(M As CodeModule)
MdnzM = M.Parent.Name
End Function

Function MdTy(M As CodeModule) As vbext_ComponentType
MdTy = M.Parent.Type
End Function

Function ShtCmpTyzM$(M As CodeModule)
ShtCmpTyzM = ShtCmpTy(CmpTyzM(M))
End Function

Function PjMd(P As VBProject, Mdn) As CodeModule
Set PjMd = P.VBComponents(Mdn).CodeModule
End Function

Function PjnzC(A As VBComponent)
PjnzC = A.Collection.Parent.Name
End Function

Function PjnzM(M As CodeModule)
PjnzM = PjnzC(M.Parent)
End Function

Function PjzM(M As CodeModule) As VBProject
Set PjzM = M.Parent.Collection.Parent
End Function

Function HasPjf(Pjf) As Boolean
HasPjf = HasPjfzV(CVbe, Pjf)
End Function

Function PjzPjfC(Pjf) As VBProject
Set PjzPjfC = PjzPjf(CVbe, Pjf)
End Function


Function Pjny() As String()
Pjny = PjnyzXls(Xls)
End Function

Sub SavCurVbe()
SavVbe CVbe
End Sub

Sub ClsMd(M As CodeModule)
M.CodePane.Window.Close
End Sub

Sub CprMd(A As CodeModule, B As CodeModule)
BrwCprDic MthlDiczM(A), MthlDiczM(B), MdDn(A), MdDn(B)
End Sub

Function CvMd(A) As CodeModule
Set CvMd = A
End Function

Sub HasCdl__Tst()
Dim Cdl$
Cdl = "Function HasCdl(M As CodeModule, Cdl$) As Boolea" & "n"
MsgBox HasCdl(CMd, Cdl)
End Sub

Function HasCdl(M As CodeModule, Cdl$) As Boolean
HasCdl = HasSubStr(Srcl(M), Cdl)
End Function

Function LinesInfzM$(M As CodeModule)
LinesInfzM = LinesInf(Srcl(M))
End Function

Function LLn(Lno, Ln) As LLn
With LLn
    .Lno = Lno
    .Ln = Ln
End With
End Function
Function AddLLn(A As LLn, B As LLn) As LLn(): PushLLn AddLLn, A: PushLLn AddLLn, B: End Function
Sub PushLLnAy(O() As LLn, A() As LLn): Dim J&: For J = 0 To LLnUB(A): PushLLn O, A(J): Next: End Sub
Sub PushLLn(O() As LLn, M As LLn): Dim N&: N = LLnSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LLnSi&(A() As LLn): On Error Resume Next: LLnSi = UBound(A) + 1: End Function
Function LLnUB&(A() As LLn): LLnUB = LLnSi(A) - 1: End Function
