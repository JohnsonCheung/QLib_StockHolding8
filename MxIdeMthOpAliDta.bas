Attribute VB_Name = "MxIdeMthOpAliDta"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Ali"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthOpAliDta."
Type AligMthDta
    LNewO As Drs
    Cln As Drs
    R123 As Drs
End Type
Function AligMthDta(Mc As Drs) As AligMthDta
Dim O As AligMthDta
Dim McCln As Drs: McCln = XMcCln(Mc)                        ' L McLin                    #Mc-Cln. ! Inl those line is &IsMthlnCxt
:                         If NoReczDrs(McCln) Then Exit Function ' Exit=>                              ! if no mth-cxt
Dim McGp  As Drs:  McGp = XMcGp(McCln)                      ' L McLin Gpno                        ! Add ^Gpno: each ^L in seq is 1-gp.  ^Gpno starts fm 1
Dim McRmk As Drs: McRmk = XMcRmk(McGp)                      ' L McLin Gpno IsColon IsRmk          ! Add ^IsRmk   wh-LTrim-FstChr-^McLin='

Dim McTRmk  As Drs:  McTRmk = XMcTRmk(McRmk)                  ' L *Rmk                        ! RmvRec wh-Mrmk.  Each gp, the above rmk lines are Mrmk, rmv them.
                                                              '                               ! [*Rmk McLin Gpno IsRmk]
Dim McInsp  As Drs:  McInsp = XMcInsp(McTRmk)                 ' L *Rmk                        ! RmvRec wh-Las-'Insp.  Each gp, the las Ln is rmk and is 'Insp, exl it.
Dim McVSfx  As Drs:  McVSfx = XMcVSfx(McInsp)                 ' L *Rmk V Sfx Rst              ! Add ^V-Sfx-Rst fm ^McLin [*Rmk McLin Gpno IsRmk]
Dim McDcl   As Drs:   McDcl = XMcDcl(McVSfx)                  ' L *Rmk V Sfx Dcl Rst          ! Add ^Dcl from ^V-Sfx
Dim McLR    As Drs:    McLR = XMcLR(McDcl)                    ' L *Rmk *V LHS RHS IsColon Rst ! Add ^LHS-RHS-IsColon fm shifting ^Rst
                                                              '                               ! ^IsColon=True when fstchr-^Rst=: and there is Only RHS
Dim McLREmp As Drs: McLREmp = XMcLREmp(McLR)                  ' L *Rmk *V LHS RHS IsColon Rst ! Set ^LHS=^V, ^RHS="X" & ^V if (^V<>"" and ^LHS="" and ^RHS=""
Dim McR123  As Drs:  McR123 = XMcR123(McLREmp)                ' L *Rmk *V *LRC R1 R2 R3       ! Add ^R1-R2-R3 from ^Rst
Dim McFill  As Drs:  McFill = XMcFill(McR123)                 ' L *Rmk *V *LRC *R *F          ! Add ^F*.  [F* F0 FSfx FRHS FR1 FR2] ^F0 is Len-of-front-spc.
Dim McAli As Drs: McAli = XMcAli(McFill)                ' L Ali                       ! Add ^Ali #Alied-Ln
Dim D1 As Drs, D2 As Drs
                         D1 = DeCeqC(McAli, "McLin Ali")  '                               ! RmvRec wh-Same-aft-align
                         D2 = SelDrs(D1, "L Ali McLin")    '                               ! Sel ^L-Aling-McLin which is Lno NewL OldL
O.LNewO = LNewO(D2.Dy)                    ' Lno NewL OldL                 ! This is req from &RplLNewO
AligMthDta = O
End Function

Function XMcTRmk(McRmk As Drs) As Drs
'Fm McRmk : L McLin Gpno IsColon IsRmk ! Add ^IsRmk   wh-LTrim-FstChr-^McLin='
'Ret      : L *Rmk                     ! RmvRec wh-Mrmk.  Each gp, the above rmk lines are Mrmk, rmv them.
'                                      ! [*Rmk McLin Gpno IsRmk] @@
Dim IGpno%, MaxGpno, A As Drs, B As Drs, O As Drs
MaxGpno = MaxEle(IntCol(McRmk, "Gpno"))
For IGpno = 1 To MaxGpno
    A = DwEq(McRmk, "Gpno", IGpno)
    B = XMcTRmkI(A)
    O = AddDrs(O, B)
Next
XMcTRmk = O
'Insp "QIde_B_AliMth.XMcTRmk", "Inspect", "Oup(XMcTRmk) McRmk", FmtDrs(XMcTRmk), FmtDrs(McRmk): Stop
End Function

Function XMcInsp(McTRmk As Drs) As Drs
'Fm McTRmk : L *Rmk ! RmvRec wh-Mrmk.  Each gp, the above rmk lines are Mrmk, rmv them.
'                   ! [*Rmk McLin Gpno IsRmk]
'Ret       : L *Rmk ! RmvRec wh-Las-'Insp.  Each gp, the las Ln is rmk and is 'Insp, exl it. @@
XMcInsp = McTRmk
If NoReczDrs(McTRmk) Then Exit Function
Dim Dr: Dr = LasEle(McTRmk.Dy)
Dim IxMcLin%: IxMcLin = IxzAy(McTRmk.Fny, "McLin")
Dim L$: L = Dr(IxMcLin)
If IsVrmkLn(L) Then
    Dim A$: A = Left(LTrim(RmvFstChr(LTrim(L))), 4)
    If A = "Insp" Then
        Pop XMcInsp.Dy
    End If
End If
'Insp "QIde_B_AliMth.XMcInsp", "Inspect", "Oup(XMcInsp) McTRmk", FmtDrs(XMcInsp), FmtDrs(McTRmk): Stop
End Function

Private Function XMcTRmkI(A As Drs) As Drs
' Fm  A :    L McLin Gpno IsRmk    #Mth-Cxt-Mrmk ! All Gpno are eq
' Ret : L McLin Gpno IsRmk ! Rmk Mrmk
Dim IxIsRmk%: AsgIx A, "IsRmk", IxIsRmk
XMcTRmkI.Fny = A.Fny
Dim J%
    Dim Dr
    For Each Dr In Itr(A.Dy)
        If Not Dr(IxIsRmk) Then GoTo Fnd 'If not a rmk-Ln, put all Ln from @J to @Oup
        J = J + 1
    Next
    Exit Function
Fnd:
    For J = J To UB(A.Dy)
        PushI XMcTRmkI.Dy, A.Dy(J)
    Next
End Function

Private Function XMcGp(McCln As Drs) As Drs
'Fm McCln : L McLin      #Mc-Cln. ! Inl those line is &IsMthlnCxt
'Ret      : L McLin Gpno          ! Add ^Gpno: each ^L in seq is 1-gp.  ^Gpno starts fm 1 @@
XMcGp = AddColzGpno(McCln, "L", "Gpno")
'Insp "QIde_B_AliMth.XMcGp", "Inspect", "Oup(XMcGp) McCln", FmtDrs(XMcGp), FmtDrs(McCln): Stop
End Function

Private Function XMcRmk(McGp As Drs) As Drs
'Fm McGp : L McLin Gpno               ! Add ^Gpno: each ^L in seq is 1-gp.  ^Gpno starts fm 1
'Ret     : L McLin Gpno IsColon IsRmk ! Add ^IsRmk   wh-LTrim-FstChr-^McLin=' @@
Dim IxMcLin%: AsgIx McGp, "McLin", IxMcLin
Dim ODy()
    Dim Dr: For Each Dr In Itr(McGp.Dy)
        PushI Dr, FstChr(LTrim(Dr(IxMcLin))) = "'"
        PushI ODy, Dr
    Next
XMcRmk = AddColzFFDy(McGp, "IsRmk", ODy)
'Insp "QIde_B_AliMth.XMcRmk", "Inspect", "Oup(XMcRmk) McGp", FmtDrs(XMcRmk), FmtDrs(McGp): Stop
End Function
