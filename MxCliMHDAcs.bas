Attribute VB_Name = "MxCliMHDAcs"
Option Compare Text
Option Explicit
Const CNs$ = "Mhd.Acs"
Const CLib$ = "QMhd."
Const CMod$ = CLib & "MxCliMHDAcs."
Dim A() As Access.Application
Function StkHld8Acs() As Access.Application:       Set StkHld8Acs = LookupAcs(StkHld8Fba): End Function
Function StkHld8TmpAcs() As Access.Application: Set StkHld8TmpAcs = LookupAcs(StkHld8TmpFba): End Function
Function DutyAcs() As Access.Application:             Set DutyAcs = LookupAcs(DutyFba): End Function
Function EStmtAcs() As Access.Application:           Set EStmtAcs = LookupAcs(EStmtFba):    End Function
Function TaxAlertAcs() As Access.Application:     Set TaxAlertAcs = LookupAcs(TaxAlertFba): End Function
Function TaxCmpAcs() As Access.Application:         Set TaxCmpAcs = LookupAcs(TaxCmpFba):   End Function

Private Function LookupAcs(Fb$) As Access.Application
With AcsIxOpt(A, Fb)
    If .Som Then
        Set LookupAcs = A(.I)
        Exit Function
    End If
End With
Push A, AcszFb(Fb)
Set LookupAcs = LasEle(A)
End Function

Private Function AcsIxOpt(A() As Access.Application, Fb$) As IntOpt
Dim Ix%, I: For Each I In Itr(A)
    If HasFb(CvAcs(I), Fb) Then AcsIxOpt = SomInt(Ix): Exit Function
    Ix = Ix + 1
Next
End Function
