Attribute VB_Name = "MxVbObjPrpDa"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbObjPrp1."

Function DrszItp(Itr, PrppSS$) As Drs
'@PrppSS:: :SS ! #Prp-Pth-Spc-Sep#
Dim Prppy$(): Prppy = SyzSS(PrppSS)
DrszItp = Drs(Prppy, DyzItrPy(Itr, Prppy))
End Function

Function DrszItrPy(Itr, Prppy$()) As Drs
DrszItrPy = Drs(Prppy, DyzItrPy(Itr, Prppy))
End Function

Function DyzItrPy(Itr, Prppy$()) As Variant()
Dim Obj As Object: For Each Obj In Itr
    Push DyzItrPy, Opvy(Obj, Prppy)
Next
End Function

Function QuietOpv(Obj, P)
On Error Resume Next
Asg Opv(Obj, P), QuietOpv
End Function

Function PrpzP1(Obj, P, Optional ThwEr As EmThw)
Select Case True
Case ThwEr = EiNoThw: Asg QuietOpv(Obj, P), PrpzP1
Case Else: Stop
End Select
End Function

Sub WAsg3PP(PP_with_NewFldEqQuoFmFld$, OPPzPrp$(), OPPzFml$(), OPPzAll$())
Dim I, S$
For Each I In SyzSS(PP_with_NewFldEqQuoFmFld)
    S = I
    If HasSubStr(S, "=") Then
        PushI OPPzAll, Bef(S, "=")
        PushI OPPzFml, I
    Else
        PushI OPPzAll, I
        PushI OPPzPrp, I
    End If
Next
End Sub

Function WFmlEr(PrpVy$(), PPzFml$()) As String()
Dim Fml, ErPmAy$(), PmAy$(), O$()
For Each Fml In Itr(PPzFml)
    PmAy = SplitComma(BetBkt(Fml))
    ErPmAy = MinusAy(PmAy, PrpVy)
    If Si(ErPmAy) > 0 Then PushI O, FmtQQ("Invalid-Pm[?] in Fml[?]", JnSpc(ErPmAy), Fml)
Next
If Si(O) > 0 Then PushI O, FmtQQ("Valid-Pm[?]", JnSpc(PrpVy))
WFmlEr = O
End Function

Private Sub DrszItrPy__Tst()
'BrwDrs DrszItpcc(Excel.Application.Addins, "Name Installed IsOpen FullName CLSId ")
'BrwDrs DrszItpcc(Fds(Db(DutyDtaFb), "Permit"), "Name Type Required")
'BrwDrs ItpDrs(Application.VBE.VBProjects, "Name Type")
BrwDrs DrszItrPy(CPj.VBComponents, SyzSS("Name Type CmpTy=ShpCmpTy(Type)"))
End Sub
