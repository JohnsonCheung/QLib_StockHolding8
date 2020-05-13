Attribute VB_Name = "MxIdeSrcDclCnstOp"
Option Explicit
Option Compare Text
Const CNs$ = "Md3Cnst"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclCnstOp."

Function EnsCLibzM(M As CodeModule, CLibv$) As Boolean
If Not IsMd(M.Parent) Then Exit Function
EnsCnst M, CLibLin(CLibv)
End Function

Sub EnsCNs(M As CodeModule, Ns$)
If Ns = "" Then
    RmvCnst M, "CNs"
Else
    EnsCnst M, CNsLin(Ns)
End If
End Sub

Sub EnsCModM()
EnsCModzM CMd
End Sub

Sub EnsCModP()
EnsCModzP CPj
End Sub

Sub EnsCModzM(M As CodeModule)
EnsCnstAft M, CModLin(M), "CLib", IsPrvOnly:=True
End Sub

Sub EnsCModzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    DoEvents
    EnsCModzM C.CodeModule
Next
End Sub

Sub SetMdNs()
'Do : Set CNsv in each module from ResFcsv("MdDrsP")
Dim D As Drs
D = ResDrs("MdDrsP")
D = SelDrs(D, "Mdn CNsv")
D = DwNBlnk(D, "CNsv")
Dim Dr: For Each Dr In Itr(D.Dy)
    Dim M As CodeModule: Set M = Md(Dr(0))
    Dim Ns$: Ns = Dr(1)
    EnsCNs M, Ns
Next
End Sub

Sub BrwMdNs()
OpnFcsv MdNsFtP
VisCXls
LasCWb.Activate
End Sub
