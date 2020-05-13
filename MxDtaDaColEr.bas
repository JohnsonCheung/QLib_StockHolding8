Attribute VB_Name = "MxDtaDaColEr"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaColEr."
Const Mtp_Col_Dup$ = "Lno(?) has Dup-?[?]"
Const Mtp_Col_NotIn$ = "Lno(?) has ?[?] which is invalid.  Valid-?=[?]"
Const Mtp_Col_NotNum$ = "Lno(?) has non-numeric-?[?]"
Const Mtp_Colx_Blnk$ = "Lno(?) has a blank [?] value"
Const Mtp_ColAy_Empty = "Lno(?) has a value of no-element-ay of a column-which-is-an-array"
Const Mtp_ColFldLikAy_NotInFny$ = "Lno(?) has FldLik[?] not in Fny[?]"
Const Mtp_ColNum_NotBet$ = "Lno(?) has ?[?] not between [?] and [?]"
Function Msgo_Col_Dup(Lnoss$, Valn$, Dup):                           Msgo_Col_Dup = FmtQQ(Mtp_Col_Dup, Lnoss, Valn, Dup):                      End Function
Function Msgo_Col_NotIn(L&, V$, Valn$, VdtValss$):                 Msgo_Col_NotIn = FmtQQ(Mtp_Col_NotIn, LnoStr(L), Valn, V, Valn, VdtValss):  End Function
Function Msgo_Col_NotNum$(L&, Valn$, V$):                         Msgo_Col_NotNum = FmtQQ(Mtp_Col_NotNum, LnoStr(L), Valn, V):                 End Function
Function Msgo_Colx_Blnk$(L&, Valn$):                               Msgo_Colx_Blnk = FmtQQ(Mtp_Colx_Blnk, LnoStr(L), Valn):                     End Function
Function Msgo_ColAy_Empty(L&):                                   Msgo_ColAy_Empty = FmtQQ(Mtp_ColAy_Empty, LnoStr(L)):                         End Function
Function Msgo_ColFldLikAy_NotInFny$(L&, F, FF$):        Msgo_ColFldLikAy_NotInFny = FmtQQ(Mtp_ColFldLikAy_NotInFny, LnoStr(L), F, FF):         End Function
Function Msgo_ColNum_NotBet(L&, Valn$, NumV, FmV, ToV):        Msgo_ColNum_NotBet = FmtQQ(Mtp_ColNum_NotBet, LnoStr(L), Valn, NumV, FmV, ToV): End Function

Function Ero_ColF_3Er(Wi_L_Colx As Drs, Fny$()) As String()
Ero_ColF_3Er = Ero_Colx_3Er(Wi_L_Colx, "F", Fny)
End Function

Function Ero_Colx_3Er(Wi_L_Colx As Drs, ColxNm$, Vy$()) As String()
Dim D As Drs: D = Wi_L_Colx
Dim A$(), B$(), C$(), VV$
VV = JnSpc(Vy)
A = Ero_Colx_NotIn(D, "F", "Fld", VV)
B = Ero_Colx_Dup(D, "F", "Fld")
C = Ero_Colx_Blnk(D, ColxNm)
Ero_Colx_3Er = AddSy(A, B)
End Function

Function Ero_Colx_Blnk(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
Dim Valn$: Valn = IIf(Valn0 = "", ColxNm, Valn0)
Dim IxL%: IxL = IxzAy(Wi_L_Colx.Fny, ColxNm)
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    If IsBlnk(Dr(IxL)) Then
        Dim L&: L = Dr(IxL)
        PushI Ero_Colx_Blnk, Msgo_Colx_Blnk(L, Valn)
    End If
Next
End Function

Function Ero_ColFldLikAy_3Er(Wi_L_LikAy As Drs, Fny$()) As String()
Dim D As Drs: D = Wi_L_LikAy
Ero_ColFldLikAy_3Er = SyzAp( _
    Ero_ColAy_Empty(D, "FldLikAy"), _
    Ero_Colx_Dup(D, "FldLikAy", "FldLik"), _
    Ero_ColFldLikAy_NotInFny(D, Fny))
End Function

Function Ero_ColAy_Empty(Wi_L_Ay As Drs, ColAyNm$) As String()
Dim IxL%, IxFny%: AsgIx Wi_L_Ay, "L Fny", IxL, IxFny
Dim Dr: For Each Dr In Itr(Wi_L_Ay.Dy)
    Dim Fny$(): Fny = Dr(IxFny)
    If Si(Fny) = 0 Then
        Dim L&: L = Dr(IxL)
        PushI Ero_ColAy_Empty, Msgo_ColAy_Empty(L)
    End If
Next
End Function

Function Ero_Colx_NotIn(Wi_L_Colx As Drs, ColxNm$, Valn$, VdtValss$) As String()
Dim IxL%, IxColx%: AsgIx Wi_L_Colx, "L " & ColxNm, IxL, 1
Dim VdtVy$(): VdtVy = SyzSS(VdtValss)
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    Dim V$: V = Dr(IxColx)
    Dim L&: L = Dr(IxL)
    If Not HasEle(VdtVy, V) Then
        PushI Ero_Colx_NotIn, Msgo_Col_NotIn(L, V, Valn, VdtValss)
    End If
Next
End Function

Function Ero_ColFldLikAy_NotInFny(Wi_L_LikAy As Drs, InFny$()) As String()
Dim IxFny%, IxL%: AsgIx Wi_L_LikAy, "L Fny", IxL, IxFny
Dim FF$: FF = JnSpc(InFny)
Dim Dr: For Each Dr In Itr(Wi_L_LikAy.Dy)
    Dim Fny$(): Fny = Dr(IxFny)
    Dim F: For Each F In Fny
        If Not HasEle(InFny, F) Then
            Dim L&: L = Dr(IxL)
            PushI Ero_ColFldLikAy_NotInFny, Msgo_ColFldLikAy_NotInFny(L, F, FF)
        End If
    Next
Next
End Function

Function Ero_Colx_Dup(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
Dim Valn$: Valn = DftStr(Valn0, ColxNm)
Dim Colx():      Colx = Col(Wi_L_Colx, ColxNm)
Dim LnoCol&(): LnoCol = LngCol(Wi_L_Colx, "L")
Dim AllLik$():          'AllLik = CvSy(AyzAyOfAy(FldLikAyCol))
Dim DupAy$():            DupAy = AwDup(AllLik)
Dim DupLik: For Each DupLik In Itr(DupAy)
    Dim Lnoss$: 'Lnoss = Lnoss_FmLnoCol_WhSyCol_HasS(LnoCol, FldLikAyCol, DupLik)
    PushI Ero_Colx_Dup, Msgo_Col_Dup(Lnoss, Valn, DupLik)
Next
If Si(Ero_Colx_Dup) > 0 Then
    Dmp Ero_Colx_Dup
    Stop
End If
End Function

Function Ero_Colx_Dup1(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
'@Valn :Nm #Val-Nm-ToBe-Shw-InMsg#
Dim Valn$: Valn = DftStr(ColxNm, Valn0)
Dim U%: U = UB(Wi_L_Colx.Dy)
Dim F$():           F = Wi_L_Colx.Fny
Dim Sy$():         Sy = StrCol(Wi_L_Colx, ColxNm)
Dim LnoCol&(): LnoCol = LngCol(Wi_L_Colx, "L")
Dim DupAy$():   DupAy = AwDup(Sy)
Dim Dup: For Each Dup In Itr(DupAy)
    Dim Lnoss$: Lnoss = Lnoss_FmLnoCol_WhStrCol_HasS(LnoCol, Sy, Dup)
    PushI Ero_Colx_Dup1, Msgo_Col_Dup(Lnoss, Valn, Dup) '<==
Next
End Function

Function Ero_Colx_NumNotBet(Wi_L_Colx As Drs, NumColxNm$, FmV, ToV) As String()
Dim IxNum%, IxL%: AsgIx Wi_L_Colx, JnSpcAp(NumColxNm, "L"), IxNum, IxL
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    Dim Num: Num = Val(Dr(IxNum))
    If Not IsBet(Num, FmV, ToV) Then
        Dim L&: L = Dr(IxL)
        PushI Ero_Colx_NumNotBet, Msgo_ColNum_NotBet(L, NumColxNm, Num, FmV, ToV)
    End If
Next
End Function

Function Ero_Colx_NotNum(Wi_L_Colx As Drs, ColxNm$) As String()
Dim IxL%, IxColxNm%: AsgIx Wi_L_Colx, "L " & ColxNm, IxL, IxColxNm
Dim Dr: For Each Dr In Wi_L_Colx.Dy
    Dim V$: V = Dr(IxColxNm)
    Dim L&
    If Not IsNumeric(V) Then
        L = Dr(IxL)
        PushI Ero_Colx_NotNum, Msgo_Col_NotNum(L, ColxNm, V)
    End If
Next
End Function
