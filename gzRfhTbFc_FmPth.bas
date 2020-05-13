Attribute VB_Name = "gzRfhTbFc_FmPth"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRfhTbFc_FmPth."
Sub RfhTbFc_FmFcIPth()
RfhTbFc_FmPth FcIPth
End Sub

Sub RfhTbFc_FmPth(Pth$)
'Aim: Create new record to table-Fc according the Fc-Import-Pth$
Dim A() As StmYM: A = StmYmAyzPth(Pth)
Dim M As StmYM
Dim J%: For J = 0 To StmYMSi(A) - 1
    M = A(J)
    Dim Sql$: Sql = "Select * from Fc" & WhereFcStm(M)
    With CurrentDb.OpenRecordset(Sql)
        If .EOF Then
            .AddNew
            !VerYY = M.Y
            !VerMM = M.M
            Dim IFx$: IFx = FcIFx(M)
            !Siz = FileLen(IFx)
            !Tim = FileDateTime(IFx)
            !DteLoad = Null
            !Stm = M.Stm
            
            .Update
        End If
    End With
Next
If IsFrmOpn("LoadFc") Then Form_LoadFc.Requery
End Sub
