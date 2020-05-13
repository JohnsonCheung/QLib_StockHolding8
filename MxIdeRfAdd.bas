Attribute VB_Name = "MxIdeRfAdd"
Option Explicit
Option Compare Text
Const CNs$ = "Pj.Rf.Op"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeRfAdd."

Sub AddRf(P As VBProject, RfLin)
Dim F$: F = RffzRfLin(RfLin)
If HasRff(P, F) Then Exit Sub
P.References.AddFromFile F
End Sub

Sub AddRfzAy(P As VBProject, RffAy$())
Dim F: For Each F In RffAy
    If Not HasRff(P, F) Then
        P.References.AddFromFile F
    End If
Next
End Sub

Sub AddRfzRff(P As VBProject, Rff$)
Const CSub$ = CMod & "AddRfzRff"
If HasRff(P, Rff) Then
    InfLn CSub, "Rff exists in Pj", "Rff Pj", Rff, P.Name
    Exit Sub
End If
P.References.AddFromFile Rff
InfLn CSub, "Rff is added to Pj", "Rff Pj", Rff, P.Name
End Sub

Sub AddRfzSrcPth(P As VBProject, SrcPth$)
AddRfzS P, RfSrczSrcPth(SrcPth)
End Sub

Sub AddRfzS(P As VBProject, RfSrc$())
Dim I: For Each I In Itr(RfSrc)
    AddRf P, I
Next
End Sub

Sub AddRfPj(P As VBProject, RfPj As VBProject)
'Do adding @RfPj to @P
Dim F$: F = Pjf(RfPj)
If HasRff(P, F) Then Exit Sub
P.References.AddFromFile F
End Sub

Sub AddStdRf(P As VBProject, StdRfn$)
Const CSub$ = CMod & "AddStdRf"
If HasRfn(P, StdRfn) Then
    Debug.Print FmtQQ("AddPjStdRf: Pj(?) already has StdRfn(?)", P.Name, StdRfn)
    Exit Sub
End If
Dim Rff$: Rff = StdRff(StdRfn)
'ChkFfnExist Rff, CSub, "StdRfFil"
P.References.AddFromFile Rff
End Sub
