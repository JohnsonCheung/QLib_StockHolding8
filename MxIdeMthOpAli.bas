Attribute VB_Name = "MxIdeMthOpAli"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CNs$ = "Mth.Op.Ali"
Const CMod$ = CLib & "MxIdeMthOpAli."

Sub AliMthzN(Mdn$, Mthn$, Optional Upd As eUpdRpt)
If Not HasMd(CPj, Mdn) Then Debug.Print "[" & Mdn & "] not exist": Exit Sub
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&: L = Mthlno(M, Mthn): If L = 0 Then Debug.Print "[" & Mthn & "] not exist": Exit Sub
AliMthzLno M, L, Upd:=Upd
End Sub

Sub AliMth(Optional Upd As eUpdRpt)
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Sub
AliMthzLno M, CMthlno, Upd:=Upd
End Sub

Sub AliMthzLno(M As CodeModule, Mthlno&, Optional Upd As eUpdRpt, Optional IsUpdSelf As Boolean)
Const CSub$ = CMod & "AliMthByLno"
Dim D1 As Drs, D2 As Drs
'== Exit if parameter error ============================================================================================
Dim Ml$:                   Ml = ContlnzM(M, Mthlno)    ' #Mth-Line#
Dim XUpd  As Boolean:    XUpd = IsUpd(Upd)              ' #Should-Upd#
Dim Mlnm$:               Mlnm = Mthn(Ml)               '  #Mth-Ln-Nm#
If IsSelfUpdEr(XUpd, IsUpdSelf, M, Mlnm) Then Exit Sub '        #Is-Self-aligning-er
Dim Mth() As LLn: Stop
'== Ali DblEqRmk (De) ================================================================================================
'   When a rmk Ln begins with '== or '-- or '.., expand it to 120 = or - or .
Dim DblChr As Drs: DblChr = DblChrRmkLNewO(Mth)                    ' L NewL OldL
:                           If XUpd Then RplLNewO M, DblChr ' <==
'== Ali Mth Cxt ======================================================================================================
Dim D As Drs: Stop
Dim Mc As AligMthDta: Mc = AligMthDta(D)
:                          If XUpd Then RplLNewO M, Mc.LNewO ' <==                           ! Upd Md-M by std Do-Lno-NewL-OldL
'== Gen Bs (Brw-Stmt) ==================================================================================================
Dim Bs      As Drs:              Bs = XBs(Mc.Cln)                               ' L BsLin ! Fst2Chr = '@
Dim Bs1     As Drs:             Bs1 = DwEqSel(Mc.R123, "IsRmk", False, "V Sfx")
Dim Bs2     As Drs:             Bs2 = DwNe(Bs1, "V", "")
Dim Vsfx    As Dictionary: Set Vsfx = DiczDrsCC(Bs2)
Dim Mdn$:                       Mdn = MdnzM(M)
Dim BsLNewO As Drs:         BsLNewO = BwSttmtLNewO(Bs, Vsfx, Mdn, Mlnm)
:                                     If XUpd Then RplLNewO M, BsLNewO

'== Crt Chd-Mth (Cm)====================================================================================================
Dim CmSel  As Drs:   CmSel = DwEq(Mc.R123, "IsRmk", False)
Dim CmV    As Drs:     CmV = SelDrs(CmSel, "V Sfx LHS RHS")    ' V Sfx LHS RHS
Dim CmDot  As Drs:   CmDot = XCmDot(CmV)                       ' V Sfx LHS RHS DotNm
Dim CmNmDD As Drs:  CmNmDD = XCmNmDD(CmDot, Mlnm)              ' V Sfx LHS RHS DotNm CmNmDD       ! DD : cm is DblDash xxx__xxx
Dim CmNmX  As Drs:   CmNmX = XCmNmX(CmNmDD)                    ' V Sfx LHS RHS DotNm CmNmDD CmNmX ! X  : cm is Xmmm
Dim CmNm   As Drs:    CmNm = XCmNm(CmNmX)                      ' V Sfx LHS RHS CmNm
Dim CmEpt  As Drs:   CmEpt = DwNe(CmNm, "CmNm", "")            ' Sfx CmNm
Dim CmEptNm$():    CmEptNm = StrCol(CmEpt, "CmNm")             ' CmNm                             ! It is ept mth ny.  They will be used create new chd mth
Dim CmActNm$():    CmActNm = MthnyzM(M)                        ' CmNm                             ! It is from chd cls of given md
Dim CmNewNm$():    CmNewNm = MinusAy(CmEptNm, CmActNm)         ' CmNm                             ! The new ChdMthny to be created.
Dim CmNew  As Drs:   CmNew = DwIn(CmEpt, "CmNm", CmNewNm)
Dim CdNewCm$:      CdNewCm = XCdNewCm(CmNew)                   '                                  ! Cd to be append to M
:                  If XUpd Then AppLines M, CdNewCm '<==

'== Upd Chd-Mth-Ln (Cml) ==============================================================================================
'   If the calling pm has been changed, the chd-mth-Ln will be updated.
Dim MlVSfx    As Drs:    MlVSfx = XMlVSfx(Ml)                             ' Ret V Sfx                           ! the Mthln's pm V Sfx
                             D1 = SelDrs(CmV, "V Sfx")
Dim CmlVSfx   As Drs:   CmlVSfx = AddDrs(MlVSfx, D1)
Dim CmlPm     As Drs:     CmlPm = XCmlPm(CmEpt)                           ' V Sfx RHS CmNm Pm
Dim CmlDclPm  As Drs:  CmlDclPm = XCmlDclPm(CmlPm, CmlVSfx)               ' V Sfx RHS CmNm Pm DclPm             ! use [CmlVSfx] & [Pm] to bld [DclPm]
Dim CmlMthRet As Drs: CmlMthRet = XCmlMthRet(CmlDclPm)                    ' V Sfx RHS CmNm Pm DclPm TyChr RetSfx
Dim CmlEpt    As Drs:    CmlEpt = XCmlEpt(CmlMthRet)                      ' V CmNm EptL
                             D1 = MthDrszM(M)                              ' L Mdy Ty Mthn Mthln
                             D1 = DwEq(D1, "Mdy", "Prv")
Dim CmlAct    As Drs:    CmlAct = SelDrsAs(D1, "L Mthn:CmNm Mthln:ActL") ' L CmNm ActL
Dim CmlJn     As Drs:     CmlJn = JnDrs(CmlEpt, CmlAct, "CmNm", "L ActL") ' V CmNm EptL L ActL                  ! som EptL & ActL may eq
                             D2 = DeCeqC(CmlJn, "EptL ActL")              ' V CmNm EptL L ActL                  ! All EptL & ActL are diff
Dim CmlLNewO  As Drs:  CmlLNewO = SelDrsAs(D2, "L EptL:NewL ActL:OldL")   ' L NewL OldL
:                                 If XUpd Then RplLNewO M, CmlLNewO        ' <==

'== Rpl Mth-Brw (Mb)====================================================================================================
'   Des: Mth-Brw is a remarked Insp-stmt in each las Ln of cm.  It insp all the inp oup
'   Lgc: Fnd-and-do MbLNewO
'        Fnd-and-do NewMb
'BrwDrs CmlEpt: Stop
Dim CmLis   As Drs:   CmLis = SelDrsAs(CmlEpt, "CmNm:Mthn EptL:Mthln") ' Mthn Mthln
Dim MbEpt   As Drs:   MbEpt = XMbEpt(CmLis, Mdn)                        ' Mthn Mthln MbStmt
Dim Cm$():               Cm = StrCol(CmLis, "Mthn")
Dim MbAct   As Drs:   MbAct = XMbAct(Cm, M)                             ' L Mthn OldL               ! OldL is MbStmt
Dim MbJn    As Drs:    MbJn = JnDrs(MbEpt, MbAct, "Mthn", "OldL L")     ' Mthn Mthln MbStmt OldL L
Dim MbSel   As Drs:   MbSel = SelDrsAs(MbJn, "L MbStmt:NewL OldL")      ' L NewL OldL
Dim MbLNewO As Drs: MbLNewO = DeCeqC(MbSel, "NewL OldL")
:                             If XUpd Then RplLNewO M, MbLNewO           ' <==

'== Crt Mth-Brw (Mb)====================================================================================================
                     D1 = LDrszJn(MbEpt, MbAct, "Mthn", "L", "HasAct") ' Mthn Mthln MbStmt L HasAct
                     D2 = DwEq(D1, "HasAct", False)                    ' Mthn Mthln MbStmt L HasAct
Dim MbNew As Drs: MbNew = SelDrsAs(D2, "Mthn MbStmt:NewL")
:                         If XUpd Then XOCrtMb M, MbNew               ' <==

'== Upd Chd-Rmk (Cr) ===================================================================================================

'-- Fnd CrEpt as Drs ---------------------------------------------------------------------------------------------------
'   Fnd #Ept      : CmNm RmkMdes  ! The expected chd mth rmk lines
'   #Fm1-McR123   : V R1 R2 R3     ! The rmk lines of each variable
'   #Fm2-CmlDclPm : V Pm           ! The v is calling chd mth is using what pm
'   #Fm3-CmlEpt   : V CmNm         ! The v is calling what chd mth

'.. Fnd #WiRmk    : V R1 R2 R3     ! all rec will have at least 1 rmk (R1..3 som not blank).............................
'   Fm  McR123
Dim CrSel   As Drs:   CrSel = SelDrs(Mc.R123, "IsColon V R1 R2 R3") ' V R1 R2 R3
Dim CrSelV  As Drs:  CrSelV = DrszFillLasIfB(CrSel, "V")           ' V R1 R2 R3                  ! Fill those blank col-V by las val
Dim FF$:                 FF = "R1 R2 R3"
Dim Sy$():               Sy = SyzAp("", "", "")
Dim CrAllL  As Drs:  CrAllL = DeVy(CrSelV, FF, Sy)                 ' IsColon V R1 R2 R3 #All-Ln ! SelRec AllLin has at least 1 rmk
Dim CrWiRmk As Drs: CrWiRmk = DwEqExl(CrAllL, "IsColon", False)    ' V R1 R2 R3                  ! SelRec ^IsColon=False

'.. Fnd #Vpr1    : V P R1 R2 R3 IsRet   ! Each V | P having what rmk.  IsRet is True....................................
'   Fm  CmlDclPm : V Pm                 ! The var calling chd mth is using what Pm
'   Fm  #WiRmk   : V R1 R2 R3           ! Each var having what Rmk
Dim CrVpm  As Drs:  CrVpm = SelDrsAs(CmlDclPm, "V Pm:P")    ' V P                ! each V is calling what Pm. Pm is SS.
Dim CrVp   As Drs:   CrVp = DrszSplitSS(CrVpm, "P")         ' V P                ! Brk P-SS into muli P
Dim CrVpr  As Drs:  CrVpr = JnDrs(CrVp, CrWiRmk, "P:V", FF) ' V P R1 R2 R3
Dim CrVpr1 As Drs: CrVpr1 = AddCol(CrVpr, "IsRet", False)   ' V P R1 R2 R3 IsRet ! All IsRet is FALSE

'.. Fnd #Ret     : V P R1 R2 R3 IsRet (P="" | IsRet=True)...............................................................
'   Fm  #WiRmk   : V R1 R2 R3
                       FF = "V P R1 R2 R3 IsRet"
Dim CrEmpP As Drs: CrEmpP = SelDrsAlwE(CrWiRmk, FF)                ' V P R1 R2 R3 IsRet ! All P & IsRet is empty
Dim CrRet  As Drs:  CrRet = DrszUpdC2VV(CrEmpP, "P IsRet", "", True) ' V P R1 R2 R3 IsRet ! All P is '' & IsRet is TRUE

'.. Fnd #RmkL : V Rmk           ! each V can map to CmNm................................................................
'   Fm  #Vrp1 : V P R1 R2 R3    ! P is pm | IsRet = false
'   Fm  #Ret :  V P R1 R2 R3    ! P is '' | IsRet = true
Dim CrMge As Drs: CrMge = AddDrs(CrVpr1, CrRet)           ' V P R1 R2 R3 IsRet ! adding CrVpr & CrRet
Dim CrAli As Drs: CrAli = AliDrs(CrMge, "V", "P R1 R2") ' V P R1 R2 IsRet    ! P R1..3 are aligned (always hav sam len)
If False Then
    Dim CrInspAli As Drs:      CrInspAli = XCrInspAli(CrAli)
    Dim Db        As Database:    Set Db = TmpDb
    CrtTblzDrs Db, "#InspAli", CrInspAli
    BrwT Db, "#InspAli"
    Stop
End If
Dim CrFst  As Drs:  CrFst = AddColzFst(CrAli, "V P") ' V P R1 R2 R3 IsRet Fst     ! P R1..3 are aligned (always hav sam len & las chr is [.]
Dim CrRmk  As Drs:  CrRmk = XCrRmk(CrFst)            ' V P R1 R2 R3 IsRet Fst Rmk ! Bld Rmk from P R1 R2 R3 & Fst
Dim CrRmkL As Drs: CrRmkL = SelDrs(CrRmk, "V Rmk")   ' V Rmk

'.. Fnd #Ept : CmNm RmkMdes :S12y ! each @CmNm should have waht @RmkMdes..............................................
'   Fm  CmlEpt:V CmNm
'   Fm  RmkL  :
Dim CrVCm    As Drs:             CrVCm = SelDrs(CmlEpt, "V CmNm") ' V CmNm
Dim CrV$():                        CrV = StrCol(CrVCm, "V")       ' V                   ! all V have chd mth
Dim CrVRmkCm As Drs:          CrVRmkCm = DwIn(CrRmkL, "V", CrV)   ' V Rmk               ! all V has chd mth
Dim CrVRmk()   As S12:           CrVRmk = S12yzDrs(CrVRmkCm)       ' V Rmk
Dim CrVRmkS()  As S12:          CrVRmkS = AddS2Sfx(CrVRmk, " @@")
Dim CrVCmD   As Dictionary: Set CrVCmD = DiczDrsCC(CrVCm)
Dim CrEpt()  As S12:            CrEpt = MapS1(CrVRmkS, CrVCmD)   ' CmNm RmkMdes ' <--

'-- Upd Chd-Rmk (Cr)----------------------------------------------------------------------------------------------------
'   If any of the calling pm's rmk is changed, the chd-mth-rmk will be updated
Dim CrAct() As S12: Stop: 'CrAct = MrmkzNy(M, Cm)
Dim CrChg() As S12: CrChg = S12yzDif(CrEpt, CrAct)         ' CmNm RmkMdes ! Only those need to change
:                          If XUpd Then XOUpdCr CrChg, M ' <==
If False Then
    Erase XX
    XBox "#1 CrEpt":     X FmtS12y(CrEpt):    XLin
    XBox "#2 CrAct":     X FmtS12y(CrAct):    XLin
    XBox "#3 CrChg":     X FmtS12y(CrChg):    XLin
    Brw XX
End If

'== Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If IsRpt(Upd) Then
    Insp CSub, "Changes", _
        "eUpdRpt DblEqRmk Ali BrwStmt " & _
        "Crt-Chd-Mth Rpl-Mth-Brw Crt-Mth-Brw Upd-Chd-Mthln Rfh-Chd-Mth-Rmk", _
        eUpdRptStr(Upd), _
        FmtDrs(DblChr), FmtDrs(Mc.LNewO), FmtDrs(BsLNewO), _
        CdNewCm, FmtDrs(MbLNewO), FmtDrs(MbNew), FmtDrs(CmlLNewO), FmtS12y(CrChg)
End If
'Insp CSub, "Cr", "CrVpr CrS12y", FmtDrs(CrVpr), FmtS12y(CrS12y): Stop
End Sub

Function XCrInspAli(CrAli As Drs) As Drs
'Fm CrAli : V P R1 R2 IsRet ! P R1..3 are aligned (always hav sam len) @@
'Fm CrAli: V P R1 R2 R3 IsRet
'Ret     : ..                 WP W1 W2 W3
If IsNeFF(CrAli, "V P R1 R2 R3 IsRet") Then Stop
Dim Dr, Dy(): For Each Dr In Itr(CrAli.Dy)
    Dim WP%: WP = Len(Dr(1))
    Dim W1%: W1 = Len(Dr(2))
    Dim W2%: W2 = Len(Dr(3))
    Dim W3%: W3 = Len(Dr(4))
    PushIAy Dr, Array(WP, W1, W2, W3)
    PushI Dy, Dr
Next
XCrInspAli = AddColzFFDy(CrAli, "WP W1 W2 W3", Dy)
'Insp "QIde_B_AliMth.XCrInspAli", "Inspect", "Oup(XCrInspAli) CrAli", FmtDrs(XCrInspAli), FmtDrs(CrAli): Stop
End Function

Function XAddColzF0(A As Drs) As Drs
Dim FstDr:        FstDr = A.Dy(0)
Dim IxMcLin%: IxMcLin = IxzAy(A.Fny, "McLin")
Dim McLin$:     McLin = FstDr(IxMcLin)
Dim T$:             T = LTrim(McLin)
Dim F0%:           F0 = Len(McLin) - Len(T)  'Len of space of fst @@McLin
                  XAddColzF0 = AddCol(A, "F0", F0)
End Function

Function XMcFill(McR123 As Drs) As Drs
'Fm McR123 : L *Rmk *V *LRC R1 R2 R3 ! Add ^R1-R2-R3 from ^Rst
'Ret       : L *Rmk *V *LRC *R *F    ! Add ^F*.  [F* F0 FSfx FRHS FR1 FR2] ^F0 is Len-of-front-spc. @@
Dim Gpno%(): Gpno = AwDis(IntCol(McR123, "Gpno"))
Dim IGpno: For Each IGpno In Itr(Gpno)
    Dim A As Drs: A = DwEq(McR123, "Gpno", IGpno) ' ..McLin.. ! Sam *Gpno
    Dim B As Drs: B = XAddColzF0(A)               ' ..F0      ! *F0 is len of spc of fst dr of @A
    Dim C As Drs: C = AddColzFiller(B, "Dcl LHS RHS R1 R2")
    Dim O As Drs: O = AddDrs(O, C)
Next
XMcFill = O
'Insp "QIde_B_AliMth.XMcFill", "Inspect", "Oup(XMcFill) McR123", FmtDrs(XMcFill), FmtDrs(McR123): Stop
End Function

Function AliRmkDblChr$(Ln)
Dim L$: L = LTrim(Ln)
If FstChr(L) = "'" Then
    Dim C2$: C2 = Fst2Chr(RmvFstChr(L))
    Select Case C2
    Case "==", "--", ".."
        AliRmkDblChr = Left(Ln, 120) & Dup(FstChr(C2), 120 - Len(Ln))
        Exit Function
    End Select
End If
AliRmkDblChr = Ln
End Function

Private Function DblChrRmkLNewO(Mth() As LLn) As Drs
'Ret   : L NewL OldL @@
Dim Dy()
Dim J%: For J = 0 To LLnUB(Mth)
    PushIOpt Dy, DrOpt_(Mth(J))
Next
DblChrRmkLNewO = LNewO(Dy)
'Insp "QIde_B_AliMth.XDeLNewO", "Inspect", "Oup(XDeLNewO) De", FmtDrs(XDeLNewO), FmtDrs(De): Stop
End Function

Private Function DrOpt_(Mth As LLn) As Opt

End Function



Function XCdNewCm$(CmNew As Drs)
'Ret :  ! Cd to be append to M @@
Dim A As Drs: A = SelDrs(CmNew, "Sfx CmNm")
If IsNeFF(CmNew, "V Sfx LHS RHS CmNm") Then Stop
Dim Dr, O$(): For Each Dr In Itr(A.Dy)
    Dim Sfx$:   Sfx = Dr(0)
    Dim CmNm$: CmNm = Dr(1)
    Dim TyChr$: TyChr = TyChrzDclSfx(Sfx)
    Dim RetSfx$: RetSfx = RetAszDclSfx(Sfx)
    PushI O, ""
    PushI O, FmtQQ("Function ??()?", CmNm, TyChr, RetSfx)
    PushI O, "End Function"
Next
XCdNewCm = JnCrLf(O)
'Insp "QIde_B_AliMth.XCdNewCm", "Inspect", "Oup(XCdNewCm) CmNew", XCdNewCm, FmtDrs(CmNew): Stop
End Function


Function XCmlEpt(CmlMthRet As Drs) As Drs
'Fm CmlMthRet : V Sfx RHS CmNm Pm DclPm TyChr RetSfx
'Ret          : V CmNm EptL
'               L Mdy Ty Mthn Mthln @@
Dim Dr, Dy(), Nm$, Ty$, Pm$, Ret$, V$, EptL$, INm%, ITy%, IPm%, IRet%, IV%
AsgIx CmlMthRet, "CmNm TyChr DclPm RetSfx V", INm, ITy, IPm, IRet, IV
'BrwDrs CmlMthRet: Stop
For Each Dr In Itr(CmlMthRet.Dy)
    Nm = Dr(INm)
    Ty = Dr(ITy)
    Pm = Dr(IPm)
    Ret = Dr(IRet)
    V = Dr(IV)
    EptL = FmtQQ("Function ??(?)?", Nm, Ty, Pm, Ret)
    PushI Dy, Array(V, Nm, EptL)
Next
XCmlEpt = DrszFF("V CmNm EptL", Dy)
'BrwDrs CmlEpt: Stop
'Insp "QIde_B_AliMth.XCmlEpt", "Inspect", "Oup(XCmlEpt) CmlMthRet", FmtDrs(XCmlEpt), FmtDrs(CmlMthRet): Stop
End Function

Function XPm$(RHS$, CmNm$)
If CmNm = "" Then Exit Function
Dim O$
If HasSubStr(RHS, "(") Then
    O = BetBkt(RHS)
Else
    O = RmvT1(RHS)
End If
XPm = JnSpc(AmTrim(SplitComma(O)))
End Function


Function XCmlPm(CmEpt As Drs) As Drs
'Fm CmEpt : Sfx CmNm
'Ret      : V Sfx RHS CmNm Pm @@
Dim IxRHS%, IxCmNm%: AsgIx CmEpt, "RHS", IxRHS, IxCmNm
Dim Dr, ODy(): For Each Dr In Itr(CmEpt.Dy)
    Dim RHS$: RHS = Dr(IxRHS)
    Dim CmNm$: CmNm = Dr(IxCmNm)
    PushI Dr, XPm(RHS, CmNm)
    PushI ODy, Dr
Next
XCmlPm = AddColzFFDy(CmEpt, "Pm", ODy)
'Insp "QIde_B_AliMth.XCmlPm", "Inspect", "Oup(XCmlPm) CmEpt", FmtDrs(XCmlPm), FmtDrs(CmEpt): Stop
End Function

Function XMcNew(Mc As Drs, McDim As Drs) As String()
'BrwDrs2 Mc, Mc, NN:="Mc McDim", Tit:="Use McDim to Upd Mc to become NewL": Stop
If JnSpc(McDim.Fny) <> "L OldL NewL" Then Stop
Dim A As Drs: A = SelDrs(McDim, "L NewL")
Dim B As Dictionary: Set B = DiczDrsCC(A)
Dim O$()
    Dim Dr, L&, Mthln$
    For Each Dr In Mc.Dy
        L = Dr(0)
        Mthln = Dr(1)
        If B.Exists(L) Then
            PushI O, B(L)
        Else
            PushI O, Mthln
        End If
    Next
XMcNew = O
End Function

Function BwSttmtLNewO(Bs As Drs, Vsfx As Dictionary, Mdn$, Mlnm$) As Drs
'Fm Bs   : L BsLin           ! Fst2Chr = '@
'Fm MlNm :         #Ml-Name. @@
Dim Dr, Dy(), S$, Ln$, L&
For Each Dr In Itr(Bs.Dy)
    L = Dr(0)
    Ln = Dr(1)
    S = WBsStmt(Ln, Vsfx, Mdn, Mlnm)
    PushI Dy, Array(L, S, Ln)
Next
BwSttmtLNewO = DrszFF("L NewL OldL", Dy)
'Insp "QIde_B_AliMth.BwSttmtLNewO", "Inspect", "Oup(BwSttmtLNewO) Bs VSfx Mdn MlNm", FmtDrs(BwSttmtLNewO), FmtDrs(Bs), VSfx, Mdn, MlNm: Stop
End Function

Function WBsStmt$(BsLin, Vsfx As Dictionary, Mdn$, Mlnm$)
Const CSub$ = CMod & "WBsStmt"
If Left(BsLin, 2) <> "'@" Then Thw CSub, "BsLin is always begin with '@", "BsLin", BsLin
Dim NN$: NN = Trim(RmvPfx(BsLin, "'@"))
Dim E$: 'E = InspEprLis(NN, Vsfx)
'WBsStmt = InspStmt(NN, E, Mdn, Mlnm)
End Function
Function XOCrtMb(M As CodeModule, MbNew As Drs)
'Fm NewMb : Cm NewMbL @@
Dim Dr: For Each Dr In Itr(MbNew.Dy)
    Dim Ln$: Ln = Dr(1)
    
    Dim Mthn$: Mthn = Dr(0)
    Dim MLno&: MLno = Mthlno(M, Mthn)
    Dim ELno&: ELno = MthEno(M, MLno)
    Dim L&:       L = ELno
    M.InsertLines L, Ln
Next
End Function

Function XMbEpt(CmLis As Drs, Mdn$) As Drs
'Fm CmLis : Mthn Mthln
'Ret      : Mthn Mthln MbStmt @@
Dim Dr, Dy()
For Each Dr In Itr(CmLis.Dy)
    Dim Mthln$: Mthln = LasEle(Dr)
    Dim MbStmt$: MbStmt = "'" & InspStmtzL(Mthln, Mdn) & ": Stop"
    PushI Dr, MbStmt
    PushI Dy, Dr
Next
XMbEpt = AddColzFFDy(CmLis, "MbStmt", Dy)
'Insp "QIde_B_AliMth.XMbEpt", "Inspect", "Oup(XMbEpt) CmLis Mdn", FmtDrs(XMbEpt), FmtDrs(CmLis), Mdn: Stop
End Function

Function XMbAct(Cm$(), M As CodeModule) As Drs
'Ret : L Mthn OldL ! OldL is MbStmt @@
Dim A As Drs: 'A = DoMthezM(M)             ' L E CmMdy Ty Mthn Mthln
Stop
Dim B As Drs: B = DwIn(A, "Mthn", Cm)
Dim Dr, Dy(): For Each Dr In Itr(B.Dy)
    Dim E&:           E = Dr(1)
    Dim L&:           L = E - 1          ' ! The Lno of MbStmt
    Dim Mthn$:     Mthn = Dr(4)
    Dim MbStmt$: MbStmt = M.Lines(L, 1)
    Select Case True
    Case HasPfx(MbStmt, "'Insp "), HasPfx(MbStmt, "Insp ")
        PushI Dy, Array(L, Mthn, MbStmt)
    End Select
Next
XMbAct = DrszFF("L Mthn OldL", Dy)
'BrwDrs MbAct: Stop
'Insp "QIde_B_AliMth.XMbAct", "Inspect", "Oup(XMbAct) Cm M", FmtDrs(XMbAct), Cm, Mdn(M): Stop
End Function

Function XBs(McCln As Drs) As Drs
'Fm McCln : L McLin #Mc-Cln. ! Inl those line is &IsMthlnCxt
'Ret      : L BsLin          ! Fst2Chr = '@ @@
Dim Dr, Dy()
For Each Dr In Itr(McCln.Dy)
    If HasPfx(Dr(1), "'@") Then PushI Dy, Dr
Next
XBs = DrszFF("L BsLin", Dy)
'Insp "QIde_B_AliMth.XBs", "Inspect", "Oup(XBs) McCln", FmtDrs(XBs), FmtDrs(McCln): Stop
End Function

Function XMcAli(McFill As Drs) As Drs
'Fm McFill : L *Rmk *V *LRC *R *F ! Add ^F*.  [F* F0 FSfx FRHS FR1 FR2] ^F0 is Len-of-front-spc.
'Ret       : L Ali              ! Add ^Ali #Alied-Ln
'                                 ! RmvRec wh-Same-aft-align
'                                 ! Sel ^L-Aling-McLin which is Lno NewL OldL @@
If NoReczDrs(McFill) Then Stop
Dim A As Drs: A = SelDrs(McFill, "L McLin Gpno Dcl IsColon LHS RHS R1 R2 R3 F0 FDcl FLHS FRHS FR1 FR2")
Dim Gpno: Gpno = DisCol(McFill, "Gpno")
Dim IGpno: For Each IGpno In Itr(Gpno)
    Dim B As Drs: B = DwEq(A, "Gpno", IGpno) ' L McLin Gpno Dcl LHS RHS R1 R2 R3 F0 FDcl FLHS FRHS FR1 FR2
    Dim C As Drs: C = XDoAli(B)
    Dim O As Drs: O = AddDrs(O, C)
Next
XMcAli = O
'Insp "QIde_B_AliMth.XMcAli", "Inspect", "Oup(XMcAli) McFill", FmtDrs(XMcAli), FmtDrs(McFill): Stop
End Function

Function XCmLin(CmDta As Drs, CmMdy$) As String()
'@CmDta  V TyChr Pm RetSfx
'Ret CmLin Mthln                     ! Mthln is always a function @@
Dim Dr, CmNm$, MthPfx$, Pm$, TyChr$, RetSfx$
For Each Dr In Itr(CmDta.Dy)
    CmNm = Dr(0)
    TyChr = Dr(1)
    Pm = Dr(2)
    RetSfx = Dr(4)
    Dim Mthln$: Mthln = FmtQQ("? Function ???(?)?", CmMdy, CmNm, TyChr, Pm, RetSfx)
    PushI XCmLin, Mthln
Next
End Function

Private Function IsSelfUpdEr(IsUpd As Boolean, IsUpdSelf As Boolean, M As CodeModule, Mlnm$) As Boolean
Const CSub$ = CMod & "XIsSelf"
'Fm MlNm :  #Ml-Name.
'Ret     :  #Is-Self-aligning-er. ! Mdn<>'QIde...' & MlNm<>'AliMthByLno @@
If Not IsUpd Then Exit Function
If IsUpdSelf Then Exit Function
Dim O As Boolean
O = Mdn(M) = "MxAliMth" And Mlnm = "AliMthByLno"
If O Then Inf CSub, "Self aligning"
IsSelfUpdEr = O
'Insp "QIde_B_AliMth.XIsSelf", "Inspect", "Oup(XIsSelf) IsUpd IsUpdSelf M MlNm", XIsSelf, IsUpd, IsUpdSelf, Mdn(M), MlNm: Stop
End Function

Function XMcCln(Mc As Drs) As Drs
'Fm Mc : L Mthln #Mth-Context.
'Ret   : L McLin  #Mc-Cln.      ! Inl those line is &IsMthlnCxt @@
Dim Dr, Dy()
For Each Dr In Itr(Mc.Dy)
    Dim Ln$: Ln = Dr(1)
    Dim IsLnMc As Boolean: IsLnMc = IsMthlnCxt(Ln) '! aa
    If IsLnMc Then PushI Dy, Dr '! aa
Next
XMcCln = DrszFF("L McLin", Dy)
'Insp "QIde_B_AliMth.XMcCln", "Inspect", "Oup(XMcCln) Mc", FmtDrs(XMcCln), FmtDrs(Mc): Stop
End Function

Sub XOUpdCr(CrChg() As S12, M As CodeModule)
Dim J&: For J = 0 To S12UB(CrChg)
    With CrChg(J)
    Stop
    'EnsMrmk M, .S1, .S2
    End With
Next
End Sub

Function XCmNmDD(CmDot As Drs, Mlnm$) As Drs
'Fm CmDot : V Sfx LHS RHS DotNm
'Fm MlNm  :                            #Ml-Name.
'Ret      : V Sfx LHS RHS DotNm CmNmDD           ! DD : cm is DblDash xxx_xxx @@
Dim MlNmDD$: MlNmDD = Mlnm & "__"
Dim Dr, Dy(): For Each Dr In Itr(CmDot.Dy)
    Dim DotNm$:       DotNm = Dr(3)
    Dim V$:               V = Dr(0)
    Dim Hit As Boolean: Hit = DotNm = MlNmDD & V
    Dim DD$:             DD = IIf(Hit, DotNm, "")
                              PushI Dr, DD
                              PushI Dy, Dr
Next
XCmNmDD = AddColzFFDy(CmDot, "CmNmDD", Dy)
'Insp "QIde_B_AliMth.XCmNmDD", "Inspect", "Oup(XCmNmDD) CmDot MlNm", FmtDrs(XCmNmDD), FmtDrs(CmDot), MlNm: Stop
End Function

Function XCmNmX(CmNmDD As Drs) As Drs
'Fm CmNmDD : V Sfx LHS RHS DotNm CmNmDD       ! DD : cm is DblDash xxx_xxx
'Ret       : V Sfx LHS RHS DotNm CmNmDD CmNmX ! X  : cm is Xmmm @@
If IsNeFF(CmNmDD, "V Sfx LHS RHS DotNm CmNmDD") Then Stop
Dim Dr, Dy(): For Each Dr In Itr(CmNmDD.Dy)
    Dim DotNm$:       DotNm = Dr(4)
    Dim V$:               V = Dr(0)
    Dim Hit As Boolean: Hit = DotNm = "X" & V
    Dim X$:               X = IIf(Hit, DotNm, "")
                              PushI Dr, X
                              PushI Dy, Dr
Next
XCmNmX = AddColzFFDy(CmNmDD, "CmNmX", Dy)
'Insp "QIde_B_AliMth.XCmNmX", "Inspect", "Oup(XCmNmX) CmNmDD", FmtDrs(XCmNmX), FmtDrs(CmNmDD): Stop
End Function

Function XCmDot(CmV As Drs) As Drs
'Fm CmV : V Sfx LHS RHS
'Ret    : V Sfx LHS RHS DotNm @@
'Fm CmV : V Sfx LHS RHS
'Ret    : V Sfx LHS RHS DotNm
If IsNeFF(CmV, "V Sfx LHS RHS") Then Stop
Dim Dr, Dy(): For Each Dr In Itr(CmV.Dy)
    Dim RHS$: RHS = Dr(3)
    PushI Dr, TakDotNm(RHS)
    PushI Dy, Dr
Next
XCmDot = AddColzFFDy(CmV, "DotNm", Dy)
'BrwDrs CmDot: Stop
'Insp "QIde_B_AliMth.XCmDot", "Inspect", "Oup(XCmDot) CmV", FmtDrs(XCmDot), FmtDrs(CmV): Stop
End Function

Function XCmNm(CmNmX As Drs) As Drs
'Fm CmNmX : V Sfx LHS RHS DotNm CmNmDD CmNmX ! X  : cm is Xmmm
'Ret      : V Sfx LHS RHS CmNm @@
'Fm CmNmX : V Sfx LHS RHS DotNm CmNmDD CmNmX
'Ret      : V Sfx LHS RHS CmNm
Dim IxSfx%, IxDD%, IxX%, IxV%, IxR%, IxL%
AsgIx CmNmX, "V Sfx LHS RHS CmNmDD CmNmX", IxV, IxSfx, IxL, IxR, IxDD, IxX
Dim Dr, Dy(): For Each Dr In Itr(CmNmX.Dy)
    Dim V$:     V = Dr(IxV)
    Dim LHS$: LHS = Dr(IxL)
    Dim RHS$: RHS = Dr(IxR)
    Dim Sfx$: Sfx = Dr(IxSfx)
    Dim DD$:   DD = Dr(IxDD)
    Dim X$:     X = Dr(IxX)
    Dim Nm$: Nm = ""
    Select Case True
    Case DD <> "": Nm = DD
    Case X <> "":  Nm = X
    End Select
    If Nm <> "" Then PushI Dy, Array(V, Sfx, LHS, RHS, Nm)
Next
XCmNm = DrszFF("V Sfx LHS RHS CmNm", Dy)
'Insp "QIde_B_AliMth.XCmNm", "Inspect", "Oup(XCmNm) CmNmX", FmtDrs(XCmNm), FmtDrs(CmNmX): Stop
End Function


Function XCmLHS(CmV As Drs) As Drs
'Fm CmV : V Sfx LHS RHS
'Ret    : V Sfx LHS RHS ! where V & ' = ' = LHS @@
Dim IxV%, IxLHS%
AsgIx CmV, "V LHS", IxV, IxLHS
Dim Dr, Dy(): For Each Dr In Itr(CmV.Dy)
    Dim V$:     V = Dr(IxV)
    Dim LHS$: LHS = Dr(IxLHS)
    If V & " = " = LHS Then Push XCmLHS.Dy, Dr
Next
XCmLHS.Fny = CmV.Fny
End Function

Function BrkRmk(Rmk$) As S3
Dim A$: A = Trim(Rmk)
If A = "" Then Exit Function
If FstChr(A) <> "'" Then Stop
Dim L$: L = Trim(RmvFstChr(A))
Dim P%: P = InStr(L, "#")
Dim E%: E = InStr(L, "!")
Dim O As S3
Select Case True
Case P = 0 And E > 0:           O.A = Bef(L, "!"): O.C = Aft(L, "!")
Case P = 0 And E = 0:           O.A = L
Case E > 0 And P > 0 And P > E: O.A = Bef(L, "!"): O.C = Aft(L, "!") 'If ! is in front of #, don't treat # of a rmk, so no-B, but only-A-and-C
Case E > 0 And P > 0:           O = BrkBet(L, "#", "!")
Case E > 0:                     O.A = Bef(L, "!"): O.C = Aft(L, "!")
Case P > 0:                     O.A = Bef(L, "#"): O.B = Aft(L, "#")
Case Else:                      O.A = L
End Select
If O.A = "" Then
    If O.B <> "" Or O.C <> "" Then
O.A = " ' "
    End If
Else
O.A = " ' " & O.A
End If
If O.B <> "" Then O.B = " #" & O.B
If O.C <> "" Then O.C = " ! " & O.C
BrkRmk = O
End Function

Function XMcR123(McLREmp As Drs) As Drs
'Fm McLREmp : L *Rmk *V LHS RHS IsColon Rst ! Set ^LHS=^V, ^RHS="X" & ^V if (^V<>"" and ^LHS="" and ^RHS=""
'Ret        : L *Rmk *V *LRC R1 R2 R3       ! Add ^R1-R2-R3 from ^Rst @@
Dim Dr: For Each Dr In Itr(McLREmp.Dy)
    Dim R As S3: R = BrkRmk(LasEle(Dr))
    Dim Dy():               PushI Dy, AddAy(AeLasEle(Dr), Array(R.A, R.B, R.C))
Next
Dim Fny$(): Fny = AddAy(AeLasEle(McLREmp.Fny), SyzSS("R1 R2 R3"))
XMcR123 = Drs(Fny, Dy)
'Insp "QIde_B_AliMth.XMcR123", "Inspect", "Oup(XMcR123) McLREmp", FmtDrs(XMcR123), FmtDrs(McLREmp): Stop
End Function
Function XDrVSfxRst(Mthln) As Variant()
Dim V$, Sfx$, Rst$
    Rst = Trim(Mthln)
    Select Case True
    Case ShfTermX(Rst, "Dim")
        V = ShfNm(Rst)
        Sfx = ShfDclSfx(Rst)
        If HasPfx(Sfx, ",") Then Stop
        Rst = Trim(RmvPfxAll(Rst, ":"))
    End Select
XDrVSfxRst = Array(V, Sfx, Rst)
End Function
Function XAvoLRqColonqRst(Ln$) As Variant()
Dim LHS$, RHS$, IsColon As Boolean, Rst$
    If FstChr(Ln) = ":" Then ' Assume the Ln with ":" is RHS only
        Dim L$: L = Trim(RmvFstChr(Ln))
        RHS = BefOrAll(L, "'")
        Dim P%: P = InStr(L, "'")
        IsColon = True
        If P > 0 Then Rst = Mid(L, P)
    Else
        Rst = Ln
        AsgAy ShfLRHS(Rst), LHS, RHS
        If Rst <> "" Then
            If Not IsVrmkLn(Rst) Then Stop
        End If
    End If
XAvoLRqColonqRst = Array(LHS, RHS, IsColon, Rst)
End Function
Function XMcLR(McDcl As Drs) As Drs
'Fm McDcl : L *Rmk V Sfx Dcl Rst          ! Add ^Dcl from ^V-Sfx
'Ret      : L *Rmk *V LHS RHS IsColon Rst ! Add ^LHS-RHS-IsColon fm shifting ^Rst
'                                         ! ^IsColon=True when fstchr-^Rst=: and there is Only RHS @@
Dim Dr: For Each Dr In Itr(McDcl.Dy)
    Dim L$:    L = Pop(Dr)
    Dim Av(): Av = XAvoLRqColonqRst(L) ' LHR RHS IsColon Rst
    Dim Dy():      PushI Dy, AddAy(Dr, Av)
Next
Dim Fny$(): Fny = AeLasEle(McDcl.Fny)
Fny = AddAy(Fny, SyzSS("LHS RHS IsColon Rst"))
XMcLR = Drs(Fny, Dy)
'BrwDrs2 McDcl, XMcLR: Stop
'Insp "QIde_B_AliMth.XMcLR", "Inspect", "Oup(XMcLR) McDcl", FmtDrs(XMcLR), FmtDrs(McDcl): Stop
End Function

Function XDclzV$(V$, WAs%, Sfx$)
If V = "" Then Exit Function
Dim O$
Select Case True
Case HasPfx(Sfx, " As "):   O = AliL(V, WAs) & Sfx
Case HasPfx(Sfx, "() As "): O = AliL(V & "()", WAs) & RmvPfx(Sfx, "()")
Case Else:                  O = V & Sfx
End Select
XDclzV = "Dim " & O & ": "
'Debug.Print XDclzV; WAs; QuoSq(Sfx); "<": Stop
End Function
Function XMcLREmp(McLR As Drs) As Drs
'Fm McLR : L *Rmk *V LHS RHS IsColon Rst ! Add ^LHS-RHS-IsColon fm shifting ^Rst
'                                        ! ^IsColon=True when fstchr-^Rst=: and there is Only RHS
'Ret     : L *Rmk *V LHS RHS IsColon Rst ! Set ^LHS=^V, ^RHS="X" & ^V if (^V<>"" and ^LHS="" and ^RHS="" @@
'Ret     : L McLin Gpno IsRmk
'          V Sfx Dcl LHS RHS Rst ! for V<>"", LHS="" and RHS="", set LHS = V and RHS = X@V
Dim IxV%, IxL%, IxR%: AsgIx McLR, "V LHS RHS", IxV, IxL, IxR
Dim Dr, Dy(): For Each Dr In Itr(McLR.Dy)
    Dim V$: V = Dr(IxV)
    Dim LHS$: LHS = Dr(IxL)
    Dim RHS$: RHS = Dr(IxR)
    If V <> "" Then
        If LHS = "" Then
            If RHS = "" Then
                Dr(IxL) = V & " = "
                Dr(IxR) = "X" & V
            End If
        End If
    End If
    PushI Dy, Dr
Next
XMcLREmp = Drs(McLR.Fny, Dy)
'Insp "QIde_B_AliMth.XMcLREmp", "Inspect", "Oup(XMcLREmp) McLR", FmtDrs(XMcLREmp), FmtDrs(McLR): Stop
End Function

Function XDcl(A As Drs) As Drs
'Fm A :      L McLin Gpno IsRmk V Sfx Rst
'Ret McDclI: L McLin Gpno IsRmk V Sfx Dcl Rst @@
Dim V$():     V = StrCol(A, "V")
Dim Sfx$(): Sfx = StrCol(A, "Sfx")
Dim WAs%:   WAs = XWAs(V, Sfx)
Dim Dr, J%, Dy(): For Each Dr In Itr(A.Dy)
    Dim Dcl$: Dcl = XDclzV(V(J), WAs, Sfx(J))
    Dim Rst$: Rst = Pop(Dr)
    PushIAy Dr, Array(Dcl, Rst)
    PushI Dy, Dr
    J = J + 1
Next
Dim Fny$(): Fny = AddAy(AeLasEle(A.Fny), Array("Dcl", "Rst"))
XDcl = Drs(Fny, Dy)
End Function
Function XDoAli(A As Drs) As Drs
'Fm A : L McLin Gpno Dcl IsColon LHS RHS R1 R2 R3 F0 FDcl FLHS FRHS FR1 FR2 ! All are in sam Gp
'Ret  : L McLin Gpno Ali
Dim L&, Gpno%, McLin$, IsColon As Boolean, Dcl$, LHS$, RHS$, R1$, R2$, R3$, F0%, FDcl%, FLHS%, FRHS%, FR1%, FR2%
If IsNeFF(A, "L McLin Gpno Dcl IsColon LHS RHS R1 R2 R3 F0 FDcl FLHS FRHS FR1 FR2") Then Stop
Dim TColon$, T0$, TDcl$, TL$, TR$, TR1$, TR2$, TR3$, Ali$, Dy()
Dim Dr: For Each Dr In Itr(A.Dy)
    AsgAy Dr, L, McLin, Gpno, Dcl, IsColon, LHS, RHS, R1, R2, R3, F0, FDcl, FLHS, FRHS, FR1, FR2
    T0 = Space(F0)
    'TColon
    'TL (TLHS)
        If IsColon Then
            TColon = ":"
            If Dcl <> "" Then Stop
            If FLHS = 0 Then
                TL = ""
            Else
                TL = Space(FLHS - 1)    'LHS should
            End If
        Else
            TColon = ""
            TL = Space(FLHS) & LHS
        End If
    TDcl = Dcl & Space(FDcl)
    TR = RHS & Space(FRHS)
    TR1 = R1 & Space(FR1)
    TR2 = R2 & Space(FR2)
    TR3 = R3
    Ali = RTrim(TColon & T0 & TDcl & TL & TR & TR1 & TR2 & TR3)
    PushI Dy, Array(L, Gpno, McLin, Ali)
Next
XDoAli = DrszFF("L Gpno McLin Ali", Dy)
'BrwDrs2 A, XDoAli: Stop
End Function


Function XMcVSfx(McInsp As Drs) As Drs
'Fm McInsp : L *Rmk           ! RmvRec wh-Las-'Insp.  Each gp, the las Ln is rmk and is 'Insp, exl it.
'Ret       : L *Rmk V Sfx Rst ! Add ^V-Sfx-Rst fm ^McLin [*Rmk McLin Gpno IsRmk] @@
Dim Dy():
    Dim IxMcLin%: AsgIx McInsp, "McLin", IxMcLin
    Dim Dr: For Each Dr In Itr(McInsp.Dy)
        Dim Av(): Av = XDrVSfxRst(Dr(IxMcLin))
        PushIAy Dr, Av
        PushI Dy, Dr
    Next
XMcVSfx = AddColzFFDy(McInsp, "V Sfx Rst", Dy)
'Insp "QIde_B_AliMth.XMcVSfx", "Inspect", "Oup(XMcVSfx) McInsp", FmtDrs(XMcVSfx), FmtDrs(McInsp): Stop
End Function

Function XMcDcl(McVSfx As Drs) As Drs
'Fm McVSfx : L *Rmk V Sfx Rst     ! Add ^V-Sfx-Rst fm ^McLin [*Rmk McLin Gpno IsRmk]
'Ret       : L *Rmk V Sfx Dcl Rst ! Add ^Dcl from ^V-Sfx @@
Dim McLin$, IxMcLin%, Dr, Dy(), IGpno
AsgIx McVSfx, "McLin", IxMcLin
For Each IGpno In AwDis(IntCol(McVSfx, "Gpno"))
    Dim A As Drs: A = DwEq(McVSfx, "Gpno", IGpno) ' L McLin Gpno IsRmk V Sfx Rst ! Sam Gpno
    Dim B As Drs: B = XDcl(A) ' L McLin Gpno IsRmk V Sfx Dcl Rst ! Adding Dcl using V Sfx
    Dim O As Drs: O = AddDrs(O, B)
Next
XMcDcl = O
'Insp "QIde_B_AliMth.XMcDcl", "Inspect", "Oup(XMcDcl) McVSfx", FmtDrs(XMcDcl), FmtDrs(McVSfx): Stop
End Function

Function XWAs%(V$(), Sfx$())
Dim C$(), J%: For J = 0 To UB(V)
    Select Case True
    Case HasPfx(Sfx(J), " As "):   Push C, V(J)
    Case HasPfx(Sfx(J), "() As "): Push C, V(J) & "()"
    End Select
Next
XWAs = WdtzAy(C)
Exit Function
Const CSub$ = CMod & "XWAs"
Thw CSub, "AAA"
End Function

Function XCmlAct(CmlEpt As Drs, M As CodeModule) As Drs
Const CSub$ = CMod & "XCmlAct"
Dim CV$(): CV = StrCol(CmlEpt, "V")
Dim Act$(): Act = StrCol(MthDrszM(M), "Mthln")
Insp CSub, "CmlAct", "CV Act", CV, Act
Stop
End Function

Function XMlVSfx(Ml$) As Drs
'Ret : Ret V Sfx ! the Mthln's pm V Sfx @@
Dim Pm$: Pm = BetBkt(Ml)
Dim P, V$, Sfx$, Dy(), L$
For Each P In Itr(SyzTrim(SplitComma(Pm)))
    L = RmvPfx(P, "ByVal ")
    L = RmvPfx(L, "Optional ")
    V = ShfNm(L)
    Sfx = L
    PushI Dy, Array(V, Sfx)
Next
XMlVSfx = DrszFF("V Sfx", Dy)
'Insp "QIde_B_AliMth.XMlVSfx", "Inspect", "Oup(XMlVSfx) Ml", FmtDrs(XMlVSfx), Ml: Stop
End Function

Function XCmlMthRet(CmlDclPm As Drs) As Drs
'Fm CmlDclPm : V Sfx RHS CmNm Pm DclPm             ! use [CmlVSfx] & [Pm] to bld [DclPm]
'Ret         : V Sfx RHS CmNm Pm DclPm TyChr RetSfx @@
Dim Dr, Sfx$, TyChr$, RetSfx$, I%, Dy()
I = IxzAy(CmlDclPm.Fny, "Sfx")
For Each Dr In Itr(CmlDclPm.Dy)
    Sfx = Dr(I)
    TyChr = TyChrzDclSfx(Sfx)
    RetSfx = RetAszDclSfx(Sfx)
    PushI Dr, TyChr
    PushI Dr, RetSfx
    PushI Dy, Dr
Next
XCmlMthRet = AddColzFFDy(CmlDclPm, "TyChr RetSfx", Dy)
'BrwDrs CmlMthRet: Stop
'Insp "QIde_B_AliMth.XCmlMthRet", "Inspect", "Oup(XCmlMthRet) CmlDclPm", FmtDrs(XCmlMthRet), FmtDrs(CmlDclPm): Stop
End Function
Function XCmlDclPm(CmlPm As Drs, CmlVSfx As Drs) As Drs
'Fm CmlPm : V Sfx RHS CmNm Pm
'Ret      : V Sfx RHS CmNm Pm DclPm ! use [CmlVSfx] & [Pm] to bld [DclPm] @@
Dim IxPm%: AsgIx CmlPm, "Pm", IxPm
Dim Dr, Dy(): For Each Dr In Itr(CmlPm.Dy)
    Dim Pm$: Pm = Dr(IxPm)
    Dim DclPm$: DclPm = XDclPm(Pm, CmlVSfx)
    PushI Dr, DclPm
    PushI Dy, Dr
Next
XCmlDclPm = AddColzFFDy(CmlPm, "DclPm", Dy)
'Insp "QIde_B_AliMth.XCmlDclPm", "Inspect", "Oup(XCmlDclPm) CmlPm CmlVSfx", FmtDrs(XCmlDclPm), FmtDrs(CmlPm), FmtDrs(CmlVSfx): Stop
End Function

Function XDclPm$(Pm$, CmlVSfx As Drs)
Const CSub$ = CMod & "XDclPm"
Dim O$(), Sfx$, P
For Each P In Itr(SyzSS(Pm))
    Sfx = VzColEq(CmlVSfx, "Sfx", "V", P)
    PushI O, P & Sfx
Next
XDclPm = JnCommaSpc(O)
'Insp CSub, "Finding DclPm(CallgPm, CmlFmMc)", "CallgPm CmlFmMc XDclPm", CallgPm, FmtDrs(CmlFmMc), XDclPm: Stop
End Function

Function XCrEpt(CrJn As Drs) As S12()
'@ CrJn : V Rmk CmNm
'Ret      : CmNm RmkMdes ! RmkMdes is find by each V in CrVpr & Mthn = V & CmPfx @@
Dim A As Drs: A = SelDrs(CrJn, "CmNm Rmk")
Dim Dr, Ly$(): For Each Dr In Itr(A.Dy)
    PushI Ly, Dr(0) & " " & Dr(1)
Next
Dim D As Dictionary: Set D = Dic(Ly, vbCrLf)
Dim O As Dictionary: Set O = AddSfxToVal(D, " @@")
XCrEpt = S12yzDic(O)
'Insp "QIde_B_MthOp.CrEpt", "Inspect", "Oup(CrEpt) CrJn", FmtS12y(CrEpt), FmtDrs(CrJn): Stop
End Function

Function XCrRmk(CrFst As Drs) As Drs
'Fm CrFst : V P R1 R2 R3 IsRet Fst     ! P R1..3 are aligned (always hav sam len & las chr is [.]
'Ret      : V P R1 R2 R3 IsRet Fst Rmk ! Bld Rmk from P R1 R2 R3 & Fst @@
If IsNeFF(CrFst, "V P R1 R2 R3 IsRet Fst") Then Stop
Dim V$, P$, R1$, R2$, R3$, IsRet As Boolean, Fst As Boolean, Rmk$
Dim Dr, Dy(): For Each Dr In Itr(CrFst.Dy)
    AsgAy Dr, V, P, R1, R2, R3, IsRet, Fst
    If R1 <> "" Then
        If Left(R1, 2) <> " '" Then Stop
        R1 = Mid(R1, 4)
    End If
    Dim Lbl$
        Select Case True
        Case Fst And IsRet
            Lbl = "'Ret" & Space(Len(P)) & " : "
        Case Fst:
            Lbl = "'Fm " & P & " : "
        Case Else
            Lbl = "'" & Space(Len(P) + 6)
        End Select
    Rmk = RTrim(Lbl & R1 & R2 & R3)
    PushI Dr, Rmk
    PushI Dy, Dr
Next
XCrRmk = AddColzFFDy(CrFst, "Rmk", Dy)
'Insp "QIde_B_AliMth.XCrRmk", "Inspect", "Oup(XCrRmk) CrFst", FmtDrs(XCrRmk), FmtDrs(CrFst): Stop
End Function

Function XCrWiRmk(CrSel As Drs) As Drs
'Fm CrSel : V R1 R2 R3
'Ret      : V R1 R2 R3 ! rmv those R1 2 3 are blank @@
Dim IxR1%, IxR2%, IxR3%
AsgIx CrSel, "R1 R2 R3", IxR1, IxR2, IxR3
Dim Dr, Dy(): For Each Dr In Itr(CrSel.Dy)
    Dim R1$: R1 = Dr(IxR1)
    Dim R2$: R2 = Dr(IxR2)
    Dim R3$: R3 = Dr(IxR3)
    If R1 <> "" Or R2 <> "" Or R3 <> "" Then
        PushI Dy, Dr
    End If
Next
XCrWiRmk = Drs(CrSel.Fny, Dy)
'Insp "QIde_B_MthOp.CrWiRmk", "Inspect", "Oup(CrWiRmk) CrSel", FmtDrs(CrWiRmk), FmtDrs(CrSel): Stop
End Function

Private Sub E(Optional Upd As eUpdRpt)
Dim Mdn$:              Mdn = "QIde_EnsCSubzMth"
Dim Mthn$:            Mthn = "EnsCSubzM"
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&:                  L = Mthlno(M, Mthn)
AliMthzLno M, L, Upd:=Upd
End Sub

Private Sub FF(Optional Upd As eUpdRpt)
Dim Mdn$: Mdn = "QIde_EnsCSubzMth"
Dim Mthn$: Mthn = "EnsCSubzM"
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&:                  L = Mthlno(M, Mthn)
AliMthzLno M, L, Upd:=Upd
End Sub
