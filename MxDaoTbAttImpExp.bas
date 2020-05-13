Attribute VB_Name = "MxDaoTbAttImpExp"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CNs$ = "Att"
Const CMod$ = CLib & "MxDaoTbAttImpExp."

Private Sub ExpAtt__Tst()
Dim T$, D As Database
T = TmpFx
ExpAttzFn D, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert HasFfn(T)
Kill T
End Sub
Function ExpCAtt$(Attn$, ToFfn$): ExpAtt CDb, Attn, ToFfn: End Function
Function ExpAtt$(D As Database, Attn$, ToFfn$)
Const CSub$ = CMod & "ExpAtt"
'Ret Exporting the first File in [Att] to [ToFfn] if Att is newer or ToFfn not exist.
'Er if no or more than one file in att, error.
'Er if any, export and return ToFfn. @@
Dim A As Attd: A = Attd(D, Attn)
FileDataFd2zAttd(A).SaveToFile ToFfn
ExpAtt = ToFfn
Inf CSub, "Att is exported", "Att ToFfn FmDb", AttnzAttd(A), ToFfn, D.Name
End Function
Function ExpAttzFn$(D As Database, Att$, AttFn$, ToFfn$)
Const CSub$ = CMod & "ExpAttzFn"
If Ext(AttFn) <> Ext(ToFfn) Then
    Thw CSub, "AttFn & ToFfn are dif extEnsion|" & _
        "To export an AttFn to ToFfn, their file extEnsion should be same", _
        "AttFn-Ext ToFfn-Ext D Attk AttFn ToFfn", _
        Ext(AttFn), Ext(ToFfn), D.Name, Att, AttFn, ToFfn
End If
If HasFfn(ToFfn) Then
    Thw CSub, "ToFfn Has, no over write", _
        "D Attk AttFn ToFfn", _
        D.Name, Att, AttFn, ToFfn
End If
Dim Fd2 As DAO.Field2
    Set Fd2 = W1F2(D, Att, AttFn$)

If IsNothing(Fd2) Then
    Thw CSub, "In record of Attk there is no given AttFn, but only Act-AttFnAy", _
        "D Given-Attk Given-AttFn Act-AttFny ToFfn", _
        D.Name, Att, AttFn, AttFnAyzNm(D, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
ExpAttzFn = ToFfn
End Function
Private Function W1F2(D As Database, Attn$, AttFn$) As DAO.Field2
With Attd(D, Attn)
    With .FldRs
        .MoveFirst
        While Not .EOF
            If !FileName = AttFn Then
                Set W1F2 = !FileData
            End If
            .MoveNext
        Wend
    End With
End With
End Function

'**ImpCAtt
Private Sub ImpAtt__Tst()
Dim T$, D As Database
T = TmpFt
WrtStr "sdfdf", T
ImpAtt D, "AA", T
Kill T
'T = TmpFt
'ExpAttToFfn "AA", T
'BrwFt T
Stop
Const Fx$ = "C:\Users\Public\Logistic\StockHolding8\WorkingDir\Templates\Stock Holding Template.xlsx"
ImpCAtt "MB52Tp", Fx
End Sub
Private Sub ImpCAtt__Tst(): ImpCAtt "Tp", StkHld8MB52Tp, "MB52Tp": End Sub
Sub ImpCAtt(Attn$, FmFfn$, Optional Attf0$): ImpAtt CDb, Attn, FmFfn, Attf0: End Sub
Sub ImpAtt(D As Database, Attn$, FmFfn$, Optional Attf0$)
Const CSub$ = CMod & "ImpAtt"
ChkFfnExist FmFfn, CSub, "Imp-To-Att-Ffn"
If Len(Attn) > 255 Then Thw CSub, "Attn-Len cannot >255", "Attn-Len Attn", Len(Attn), Attn
Dim Attf$: Attf = DftStr(Attf0, Fn(FmFfn))
Dim A As Attd: A = Attd(D, Attn)
If W2HasAttf(A, Attf) Then
    W2Rpl A, Attf, FmFfn
Else
    W2Imp A, Attf, FmFfn
End If
W2UpdTbAttd D, Attn, Attf, FmFfn
End Sub
Private Function W2HasAttf(A As Attd, Attf) As Boolean: W2HasAttf = HasRecRsFeq(A.FldRs, "FileName", Attf): End Function
Private Sub W2Imp(A As Attd, Attf$, Ffn$)
A.RecRs.Edit
A.FldRs.AddNew
CvFd2(A.FldRs!FileData).LoadFromFile W2RenTo(Ffn, Attf)
W2RenBack Ffn, Attf
A.FldRs.Update
A.RecRs.Update
End Sub
Private Function W2RenTo$(Ffn$, Attf$)
End Function
Private Sub W2RenBack(Ffn$, Attf$)

End Sub

Private Sub W2Rpl(A As Attd, Attf$, Ffn$)
A.RecRs.Edit
W2Fd2(A, Attf).LoadFromFile Ffn
A.FldRs.Update
A.RecRs.Update
End Sub
Private Function W2Fd2(A As Attd, Attf$) As DAO.Field2
With A.FldRs
    .MoveFirst
    While Not .EOF
        If !FileName = Attf Then
            .Edit
            Set W2Fd2 = !FileData
            Exit Function
        End If
        .MoveNext
    Wend
End With
Imposs "W2Fd2"
End Function
Private Sub W2UpdTbAttd(D As Database, Attn$, Attf$, FmFfn$)
With W2IupTbAttd(D, Attn, Attf)
!FilTim = DtezFfn(FmFfn)
!FilSi = SizFfn(FmFfn)
.Update
End With
End Sub
Private Function W2IupTbAttd(D As Database, Attn, Attf) As DAO.Recordset: Set W2IupTbAttd = RszSkvap(D, "Attd", AttId(D, Attn), Attf): End Function
