Attribute VB_Name = "MxIdeSrcCSubDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CNs$ = "CSub"
Const CMod$ = CLib & "MxIdeSrcCSubDrs."
Public Const CSubFF$ = "Mdn L Mthl CurCSubLno CurCSubLin EptCSubLno EptCSubLin OpLno LinOp OldL NewL" ' FoCSub.LinOp = [Rpl | Ins | Dlt]

Private Function IsEroDltOrRplLNewO(MdnDryLn) As Boolean
Const CSub$ = CMod & "IsEroDltOrRplLNewO"
Dim I As MdLnOpRec: I = MdLnOpRec(MdnDryLn)
Dim M As CodeModule: Set M = Md(I.Mdn)
Dim NowL$: NowL = M.Lines(I.OpLno, 1)
Select Case I.LinOp
Case "Rpl", "Dlt": If NowL <> I.OldL Then IsEroDltOrRplLNewO = True: InfLn CSub, "Dlt/Rpl Ln not expect", "LinOp NowL OldL", I.LinOp, NowL, I.OldL
End Select
End Function

Private Sub DoCSubzP__Tst()
Dim D As Drs
GoSub T1
Exit Sub
T1: D = CSubDrsP: GoTo Tst
YY: BrwDrsN CSubDrsP: Return ', "Mdn L Lno Mthl LinOp OldL NewL")
Tst:
    'Using D :: :Wi Mthl LinOp Lno OldL NewL:
    Dim IsEr As Boolean
        IsEr = False
        Dim Dr: For Each Dr In D.Dy
            If IsEroDltOrRplLNewO(Dr) Then IsEr = True
        Next
    If IsEr Then Stop
    Return
End Sub

Private Sub DoCSubzM__Tst()
Dim D As Drs
GoSub T1
Exit Sub
T1: D = CSubDrszM(Md("MxCmpOp")): GoTo Tst
Tst:
    Dim IsEr As Boolean
        IsEr = False
        Dim Dr: For Each Dr In D.Dy
            If IsEroDltOrRplLNewO(Dr) Then IsEr = True
        Next
    If IsEr Then Stop
    Return
End Sub

Function CSubDrsP() As Drs
CSubDrsP = CSubDrszP(CPj)
End Function

Private Function CSubDrszP(P As VBProject) As Drs
CSubDrszP = CSubDrs(MthlDrszP(P))
End Function

Private Function CSubDrszM(M As CodeModule) As Drs
CSubDrszM = CSubDrs(MthlDrszM(M))
End Function

Private Function CSubDyzMth(Mdn$, L&, Mthl$) As Variant()
'Ret : :Dy ! #One-Mth-of-DyoCSub# 0,1 or 2 Dr.  In some case it will return 2 Dr: Dlt/Ins
Dim Cur As LLn
Dim Ept As LLn
    Dim A_MthLy$(): A_MthLy = SplitCrLf(Mthl)
                              Cur = CCSubLLn(L, A_MthLy)
                              Ept = EptCSubLLn(L, A_MthLy)
'Dim DyoCSubzMth1(): Dim D As Drs, Inf As LLn, Ept As LLn
    Dim B_Mdy(): B_Mdy = MdyLnDy(Cur, Ept) ' :Dyo :
    If Si(B_Mdy) > 0 Then
        Dim B_Hdr(): B_Hdr = Av(Mdn, L, Mthl, Cur.Lno, Cur.Ln, Ept.Lno, Ept.Ln)
        Dim B_Dr: For Each B_Dr In B_Mdy
            PushI CSubDyzMth, AddAy(B_Hdr, B_Dr)
        Next
    End If
End Function

Function CSubDrs(MthlDrs As Drs) As Drs
Dim ODy()
    Dim Dr: For Each Dr In Itr(MthlDrs.Dy)
        Dim Mdn$:   Mdn = Dr(0)
        Dim L&:       L = Dr(1)
        Dim Mthl$: Mthl = Dr(2)
        PushIAy ODy, CSubDyzMth(Mdn, L, Mthl)
    Next
CSubDrs = DrszFF(CSubFF, ODy)
'Insp CSub, "The Dlt-LinOp-OldL is matching the Md-Line?", "DoCSub MthlDrs", FmtDrs(DoCSub), FmtDrs(MthlDrs): Stop
End Function


Private Function Y_Mthl$()
Const CSub$ = CMod & "Y_Mthl"
Y_Mthl$ = _
"Function XWAs%(V$(), Sfx$())" & vbCrLf & _
"Dim C$(), J%: For J = 0 To UB(V)" & vbCrLf & _
"   Select Case True" & vbCrLf & _
"   Case HasPfx(Sfx(J), "" As ""):   Push C, V(J)" & vbCrLf & _
"   Case HasPfx(Sfx(J), ""() As ""): Push C, V(J) & ""()""" & vbCrLf & _
"   End Select" & vbCrLf & _
"Next" & vbCrLf & _
"XWAs = WdtzAy(C)" & vbCrLf & _
"Exit Function" & vbCrLf & _
"Const CSub$ = CMod & ""XWAs""" & vbCrLf & _
"Thw CSub, ""AAA""" & vbCrLf & _
"End Function"
End Function

Private Function Y_MthlDrs() As Drs
Dim Dr(): Dr = Array("MxAliMth", 875, "XWAs", Y_Mthl)
Y_MthlDrs = DrszFF("Mdn L Mthn Mthl", CvAv(Array((Dr))))
End Function

Private Sub CSubDrs__Tst()
Const CSub$ = CMod & "Z_CSubDrs"
Dim Mthnc As Drs
GoSub T1
Exit Sub
T1:
    BrwDrsN Y_MthlDrs: Stop
    '-- ---------- --- ---- ------------------------------------------------------ --------------------------- ------- --------------------------- ----------
    'Ix Mdn        L   Mthn Mthl                                                   CSubLin                     CSubLno EptCSubLin                  EptCSubLno
    '-- ---------- --- ---- ------------------------------------------------------ --------------------------- ------- --------------------------- ----------
    '1  MxAliMth 875 XWAs Function XWAs%(V$(), Sfx$())                           Const CSub$ = CMod & "XWAs" 884     Const CSub$ = CMod & "XWAs" 876
    '                       Dim C$(), J%: For J = 0 To UB(V)
    '                           Select Case True
    '                           Case HasPfx(Sfx(J), " As "):   Push C, V(J)
    '                           Case HasPfx(Sfx(J), "() As "): Push C, V(J) & "()"
    '                           End Select
    '                       Next
    '                       XWAs = WdtzAy(C)
    '                       Exit Function
    '                       Const CSub$ = CMod & "XWAs"
    '                       Thw CSub, "AAA"
    '                       End Function
    '-- ---------- --- ---- ------------------------------------------------------ --------------------------- ------- --------------------------- ----------
    GoTo Tst
Tst:
    Dim Act As Drs
    Act = CSubDrs(Y_MthlDrs)
    Brw FmtDrs(Act)
    Stop
    Return
End Sub


Private Function EptCModLin$(Mdn)
EptCModLin = FmtQQ("Const CMod$ = ""?.""", Mdn)
End Function
