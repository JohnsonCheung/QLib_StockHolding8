Attribute VB_Name = "MxIdeMthDic"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthDic."
':SDiMdnqSrc$ = "SDiMdnqSrc is Srtd-Mdn-To-SrcL."

Function MthlDizP(P As VBProject) As Dictionary
Dim C As VBComponent
For Each C In P.VBComponents
    PushDic MthlDizP, MthlDiczM(C.CodeModule)
Next
End Function

Private Sub MthlDizP__Tst()
Dim A As Dictionary: Set A = MthlDizP(CPj)
Ass IsDiiLines(A) '
Vc A
End Sub

Private Sub MthlDicM__Tst()
B MthlDicM
End Sub

Function MthlDicP() As Dictionary
Set MthlDicP = MthlDizP(CPj)
End Function

Function MthlDicM() As Dictionary
Set MthlDicM = MthlDiczM(CMd)
End Function

Function MthlDic(Src$()) As Dictionary 'Key is MthDn, Val is MthLWiMrmk
':MthlDi: :MthKn-Mthl-Di
Set MthlDic = New Dictionary
With MthlDic
    .Add "*Dcl", DcllzSrc(Src)
    Dim Ix: For Each Ix In MthixItr(Src)
        .Add MthKnzL(Src(Ix)), Mthl(Src, Ix)
    Next
End With
End Function

Function MthlDiczM(M As CodeModule) As Dictionary
Set MthlDiczM = MthlDic(Src(M))
End Function

Function SrtdMthlDiczP(P As VBProject) As Dictionary
Set SrtdMthlDiczP = SrtDic(MthlDizP(P))
End Function

Function SrtdMthlDicP() As Dictionary
Set SrtdMthlDicP = SrtdMthlDiczP(CPj)
End Function

Sub BrwSrtRptzM(M As CodeModule)
Dim Old$: Old = Srcl(M)
Dim NewLines$: NewLines = SrtdSrclzM(M)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print Mdn(M), O
End Sub


Sub SrtPj(P As VBProject)
BkuFfn Pjf(P)
Dim C As VBComponent
For Each C In P.VBComponents
    SrtMd C.CodeModule
Next
End Sub


Private Sub Dcl_BefAndAft_Srt__Tst()
Const Mdn$ = "DqStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = Src(Md(Mdn))
B = SrtSrc(A)
A1 = DclzSrc(A)
B1 = DclzSrc(B)
Stop
End Sub
