Attribute VB_Name = "MxDaoDbCrtTbl"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoDbCrtTbl."

Sub CrtTblzEmpFm(D As Database, T, FmTbl$)
RunQ D, SqlSelStar_Into_Fm_WhFalse(T, FmTbl)
End Sub

Sub CrtTblzDrs(D As Database, T, Drs As Drs)
CrtEmpTblzDrs D, T, Drs
InsTblzDy D, T, Drs.Dy
End Sub

Sub CrtTblzDup(D As Database, T, FmTbl, KK$)
Dim K$, Jn$, Tmp$, J%
Tmp = "##" & TmpNm
K = QpFis(KK)
Dim Into$
D.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, FmTbl, K)
D.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", Into, FmTbl, Tmp, Jn)
DrpT D, Tmp
End Sub

Sub CrtTblzJnFld(D As Database, T, KK$, JnFld$, Optional Sep$ = " ")
Dim Tar$, LisFld$
    Tar = T & "_Jn_" & JnFld
    LisFld = JnFld & "_Jn"
Stop 'RunQ D, SqlSel_Fny_Into_Fm(Ny(KK), Tar, T)
AddFld D, T, LisFld, dbMemo
Dim KKIdx&(), JnFldIx&
    KKIdx = Ixy(Fny(D, T), Ny(KK))
    JnFldIx = IxzF(D, T, JnFld)
InsTblzDy D, T, DyJnFldKK(DyzT(D, T), KKIdx, JnFldIx)
End Sub
