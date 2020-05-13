Attribute VB_Name = "MxXlsLofEr1"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxXlsLofEr1."

Private Sub ErzErSrc__Tst()
Dim Src$(), Mdn, ErNy$(), ErMthAet As Dictionary
'GoSub T1
'GoSub YY1
GoSub YY2
Exit Sub
T1:

YY2:
    Src = SrczMdn("MXls_Lof_EoLof")
    GoSub Tst
    Brw Act
    Return
YY1:
    GoSub Set_Src
    Mdn = "XX"
    GoSub Tst
    Brw Act
    Return
Tst:
    Act = ErzErSrc(Src, ErMthAet, Mdn)
    Return
Set_Src:
    Const X$ = "'GenErMsg-Src-Beg." & _
    "|'Val_NotNum      Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number" & _
    "|'Val_NBet      Lno#{Lno} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})" & _
    "|'Val_NotInLis    Lno#{Lno} is [{T1$}] line having invalid Val({ErVal$}).  See valid-value-{VdtValNm$}" & _
    "|'Val_FmlFld      Lno#{Lno} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErFny$()}).  Valid-Fny are [{VdtFny$()}]" & _
    "|'Val_FmlNotBegEq Lno#{Lno} is [Fml] line having [{Fml$}] which is not started with [=]" & _
    "|'Fld_NotInFny    Lno#{Lno} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]" & _
    "|'Fld_Dup         Lno#{Lno} is [{T1$}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno}" & _
    "|'Fldss_NotSel    Lno#{Lno} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]" & _
    "|'Fldss_DupSel    Lno#{Lno} is [{T1$}] line having" & _
    "|'Lon            Lno#{Lno} is [Lo-Nm] line having value({Val$}) which is not a good name" & _
    "|'Lon_Mis        [Lo-Nm] line is missing" & _
    "|'Lon_Dup        Lno#{Lno} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno}" & _
    "|'Tot_DupSel      Lno#{Lno} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored." & _
    "|'Bet_N3Fld        Lno#{Lno} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]" & _
    "|'Bet_EqFmTo      Lno#{Lno} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal." & _
    "|'Bet_FldSeq      Lno#{Lno} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]" & _
    "|'GenErMsg-Src-End." & _
    "|Const M_Bet_FldSeq$ = 1"
    Src = SplitVBar(X)
    Return
End Sub

Function ContBei(Src$(), Ix&) As Bei
ContBei.Bix = Ix
ContBei.Eix = ContEIx(Src, Ix)
End Function

Function ErzErSrc(ErSrc$(), ErMthnSet As Dictionary, Optional Mdn = "?") As String()
Const CSub$ = CMod & "ErzErSrc"
'Init ErSrc
If ErMthnSet.IsEmp Then Inf CSub, "No GenErMsg-ErSrc-Beg. / GenErMsg-ErSrc-End.", "Md", Mdn: Exit Function
Dim O$(), O1$(), O2$()
'O1 = SrcRplConstDic(ErSrc, ErConstDic(ErMthnSet)): 'Brw O1: Stop
'O2 = RmvSrcMth(O1, ErMthnSet):       'Brw LyzNNAp("MthToRmv BefDltMth AftDltMth", ErMthnSet, O1, O2): Stop
'O = AddSy(O2, ErMthlny):            'Brw O:Stop
ErzErSrc = O
End Function

Function SrcRplConstDic(Src$(), ConstDic As Dictionary) As String()
Dim Cnstn, Dcl$(), Bdy$(), Dcl1$(), Dcl2$()
AsgDclAndBdy Src, Dcl, Bdy
Dcl1 = DclRmvCnstnSet(Dcl, KeyAet(ConstDic)): 'Brw Dcl1: Stop
'Brw LyzLinesDicItems(ConstDic): Stop
Dcl2 = Sy(Dcl1, LyzLinesDicItems(ConstDic), Bdy): 'Brw Dcl2: Stop
SrcRplConstDic = Dcl2
End Function

Function DclRmvCnstnSet(Dcl$(), CnstnSet As Dictionary) As String()
Dim L: For Each L In Itr(Dcl)
    If Not CnstnSet.Exists(Cnstn(L)) Then
        PushI DclRmvCnstnSet, L
    End If
Next
End Function

Sub AsgDclAndBdy(Src$(), ODcl$(), OBdy$())
Dim J&, F&, U&
U = UB(Src)
F = FstMthix(Src)
If F < 0 Then
    Erase OBdy
    ODcl = Src
    Exit Sub
End If
For J = 0 To F - 1
    PushI ODcl, Src(J)
Next
For J = F To U
    PushI OBdy, Src(J)
Next
End Sub

Function ErMthnSet(ErMthny$()) As Dictionary
Set ErMthnSet = Aet(ErMthny)
End Function

Function ErMthny(ErNy$()) As String()
Dim I
For Each I In Itr(ErNy)
'    PushI ErMthny, ErMthn(I)
Next
End Function

Function ErConstDic(ErMthnSet As Dictionary, ErMsgAy$()) As Dictionary
Const C$ = "Const  M_?$ = ""?"""
Set ErConstDic = New Dictionary
Dim ErNm, J%
For Each ErNm In ErMthnSet.Itms
    ErConstDic.Add ErCnstn(ErNm), FmtQQ(C, ErNm, ErMsgAy(J))
    J = J + 1
Next
End Function

Function ErCnstn$(ErNm)
ErCnstn = "M_" & ErNm
End Function

Private Sub ErMthlny__Tst()
Dim ErNy$(), ErMsgAy$(), ErMthlny$()
'GoSub Z
GoSub T1
Exit Sub
Z:
    Brw ErMthlny
    Return
T1:
    ErNy = Sy("Val_NotNum")
    ErMsgAy = Sy("Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number")
    Ept = Sy("Function MsgzVal_NotNum(Lno, T1, Val$) As String(): MsgzVal_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val): End Function")
    GoTo Tst
Tst:
    Act = ErMthlny
    C
    Return
End Sub

Private Sub Init__Tst()
'Init Src(Md("MXls_Lof_EoLof"))
'If Not HasEle(ErNy, "Bet_FldSeq") Then Stop
'Bet_FldSeq
Stop
End Sub

Function ErMthlny(ErNy$(), ErMsgAy$()) As String() 'One ErMth is one-MulStmtLin
Dim J%, O$()
For J = 0 To UB(ErNy)
'    PushI O, ErMthLByNm(ErNy(J), MsgAy(J))
Next
'ErMthlny = FmtMulStmtSrc(O)
End Function

Function ErMthLByNm$(ErNm$, ErMsg$)
Dim CNm$:         CNm = ErCnstn(ErNm)
Dim ErNy$():     ErNy = MacroNy(ErMsg)
Dim Pm$:           Pm = JnCommaSpc(AwDis(ErNy))
Dim Calling$: Calling = Jn(AmAddPfx(DclNy(ErNy), ", "))
Dim Mthn:     'Mthn = ErMthn(ErNm)
ErMthLByNm = FmtQQ("Function ?(?) As String():? = FmtMacro(??):End Function", _
    Mthn, Pm, Mthn, CNm, Calling)
End Function
