Attribute VB_Name = "MxIdePjInf"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdePjInf."

Function XlsOpnXla(X As Excel.Application, Fxa) As VBProject
'Ret: Ret :Pj of @Fxa fm @X if exist, else @X.Opn @Fxa
Dim O As VBProject: Set O = PjzPjf(X.Vbe, Fxa)
If Not IsNothing(O) Then Set XlsOpnXla = O: Exit Function
Set XlsOpnXla = OpnXFx(X, Fxa).VBProject
End Function

Function PjzFxa(Fxa) As VBProject
'Ret: Ret :Pj of @Fxa fm @Xls if exist, else @Xls.Opn @Fxa
Set PjzFxa = XlsOpnXla(Xls, Fxa)
End Function

Function HasFxa(Fxa$) As Boolean
HasFxa = HasStrEle(PjfnAyV, Fn(Fxa))
End Function

Sub OpnFxa(Fxa$)
Const CSub$ = CMod & "OpnFxa"
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then
    Inf CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, PjnyV
    Exit Sub
End If
Xls.Workbooks.Open Fxa
End Sub

Function PjnzFxa$(Fxa)
PjnzFxa = Fnn(RmvNxtNo(Fxa))
End Function

Sub CrtFxa(Fxa$)
Const CSub$ = CMod & "CrtFxa"
'Do: crt an emp Fxa with pjn derived from @Fxa
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then Thw CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, PjnyV
Dim Wb As Workbook: Set Wb = Xls.Workbooks.Add
Wb.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn 'Must save first, otherwise PjzFxa will fail.
PjzFxa(Fxa).Name = PjnzFxa(Fxa)
Wb.Close True
End Sub

Function FrmFfnAy(Pth) As String()
Dim I: For Each I In Itr(Ffny(Pth, "*.frm.txt"))
    PushI FrmFfnAy, I
Next
End Function

Function ClsAyzP(P As VBProject) As CodeModule()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCls(C) Then
        PushObj ClsAyzP, C
    End If
Next
End Function

Function ClsNy(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsCls(C) Then
        PushI ClsNy, C.Name
    End If
Next
End Function

Private Sub CmpAyzP__Tst()
Dim Act() As VBComponent
Dim C, T As vbext_ComponentType
For Each C In CmpAyzP(CPj)
    T = CvCmp(C).Type
    If T <> vbext_ct_StdModule And T <> vbext_ct_ClassModule Then Stop
Next
End Sub

Function CmpAyzP(P As VBProject) As VBComponent()
If IsProtectvvInf(P) Then Exit Function
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMd(C) Then
        PushObj CmpAyzP, C
    End If
Next
End Function

Function IsPjNoClsNoMod(P As VBProject) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
IsPjNoClsNoMod = True
End Function

Function ModItrzP(P As VBProject)
Asg Itr(ModAyzP(P)), _
    ModItrzP
End Function

Function ModAyzP(P As VBProject) As CodeModule()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent: For Each C In P.VBComponents
    If C.Type = vbext_ct_StdModule Then
        PushObj ModAyzP, C.CodeModule
    End If
Next
End Function

Function NoNsMdNyP() As String()
NoNsMdNyP = NoNsMdNyzP(CPj)
End Function

Function NoNsMdNyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If IsNoNsMd(C.CodeModule) Then
        PushI NoNsMdNyzP, C.Name
    End If
Next
End Function

Function IsNoNsMd(M As CodeModule)
IsNoNsMd = CNsv(Dcl(M)) = ""
End Function

Function CMdNy() As String()
CMdNy = Mdny(CPj)
End Function

Function MdNyWiPrpV() As String()
MdNyWiPrpV = MdNyWiPrpzV(CVbe)
End Function

Function MdNyWiPrpzV(A As Vbe) As String()
Dim Mdn, I
For Each I In MdNyzV(A)
    Mdn = I
    If IsMdnWiPrp(Mdn) Then
        PushI MdNyWiPrpzV, Mdn
    End If
Next
End Function

Function IsMdnWiPrp(Mdn) As Boolean
Dim M As CodeModule: Set M = Md(Mdn)
Dim J&
For J = 1 To M.CountOfLines
    If IsPrpln(M.Lines(J, 1)) Then IsMdnWiPrp = True: Exit Function
Next
End Function

Function MdNyV() As String()
MdNyV = MdNyzV(CVbe)
End Function

Function MdAyzNy(Mdny$()) As CodeModule()
Dim P As VBProject: Set P = CPj
Dim N: For Each N In Itr(Mdny)
    PushI MdAyzNy, MdzP(P, N)
Next
End Function

Function MdAyzPubMthn(PubMthn) As CodeModule()
MdAyzPubMthn = MdAyzNy(MdNyzPubMthn(PubMthn))
End Function

Function MdNyzPubMthn(PubMthn) As String()
MdNyzPubMthn = MdnAetzPubMthn(PubMthn).Sy
End Function

Function MdnAetzPubMthn(PubMthn) As Dictionary
'Set MdnAetzPubMthn = MthnqMdnRelV.ParChd(PubMthn)
End Function

Function MdnAetzM(Mthn) As Dictionary
Set MdnAetzM = MthnRelMdnP.ParChd(Mthn)
End Function

Function PubMthnqMdnRelV() As Dictionary
'Set PubMthnqMdnRelV = PubMthnqMdnRelzV(CVbe)
End Function

Function CmpAyP() As VBComponent()
CmpAyP = CmpAyzP(CPj)
End Function

Function MdAy() As CodeModule()
MdAy = MdAyzP(CPj)
End Function

Function CmpItr(P As VBProject)
Asg CmpItr, _
    Itr(CmpAyzP(P))
End Function

Function MdItr(P As VBProject)
Asg MdItr, _
    Itr(MdAyzP(P))
End Function

Function MdAyzP(P As VBProject) As CodeModule()
MdAyzP = MdAyzC(CmpAyzP(P))
End Function


Function DftPj(P As VBProject) As VBProject
If IsNothing(P) Then
    Set DftPj = CPj
Else
    Set DftPj = P
End If
End Function
