Attribute VB_Name = "MxIdeSrcSngQExmRmk"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcSngQExmRmk."
Function RmkzSngQExmLin$(Ln)

End Function
Function RmkzTyDfnRmkLy$(TyDfnRmkLy$())
Dim R$, O$()
Dim L: For Each L In Itr(TyDfnRmkLy)
    If FstChr(L) = "'" Then
        Dim A$: A = LTrim(RmvFstChr(L))
        If FstChr(A) = "!" Then
            PushNB O, LTrim(RmvFstChr(A))
        End If
    End If
Next
RmkzTyDfnRmkLy = JnCrLf(O)
End Function
Function SngQExmRe() As RegExp
Static O As RegExp
If IsNothing(O) Then
End If
End Function

Function IsLnSngQExm(L) As Boolean
IsLnSngQExm = SngQExmRe.Test(L)
End Function
