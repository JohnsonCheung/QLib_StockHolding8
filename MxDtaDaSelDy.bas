Attribute VB_Name = "MxDtaDaSelDy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaSelDy."

Function DywKeyDr(Dy(), KeyDr(), KeyDy()) As Variant()
DywKeyDr = AwIxy(Dy, SubIxy_(KeyDr, KeyDy))
End Function

Private Function SubIxy_(ByDr(), InDy()) As Long()
Dim Dr, I&: For Each Dr In Itr(InDy)
    If IsEqDr(Dr, ByDr) Then
        #If False Then
        Stop
        Debug.Print I
        Debug.Print JnSpc(Dr)
        Debug.Print JnSpc(ByDr)
        Debug.Print
        #End If
        PushI SubIxy_, I
    End If
    I = I + 1
Next
End Function

Private Sub SubIxy___Tst()
Dim KeyDr()
Dim KeyDy()
    KeyDy = SelDrs(MdDrsP, "CLibv CNsv").Dy
Dim SubIxy1&()
    KeyDr = Array("QGit", Empty)
    SubIxy1 = SubIxy_(KeyDr, KeyDy)
Dim SubIxy2&()
    KeyDr = Array("QAct", Empty)
    SubIxy2 = SubIxy_(KeyDr, KeyDy)
Dmp SubIxy1
Debug.Print
Dmp SubIxy2

End Sub
