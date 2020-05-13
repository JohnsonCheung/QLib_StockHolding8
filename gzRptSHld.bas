Attribute VB_Name = "gzRptSHld"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRptSHld."
Option Base 0
Public Const SH10ValFF$ = "ScCsg HkdCsg ScDf HkdDf ScDp HkdDp ScGit HkdGit ScTot HkdTot"
Public Const SH4KpiFF$ = "StkDays StkMths RemSC TarStkMths"
Public Const SH14ValFF$ = "ScCsg HkdCsg ScDf HkdDf ScDp HkdDp F1 ScGit HkdGit F2 ScTot HkdTot"

'Sub T_OFc():      UUOupFc Samp86Dec24:    End Sub
'Sub T_OStkHld():  UUOupStkHld__Tst: End Sub
'Sub T_OStkDays(): UUOupStkDays Samp86Dec24: End Sub
'Sub T_RptShld():  RptShld__Tst:   End Sub
'Sub T_UUGen(): UUGen__Tst: End Sub
Private Sub RptShld__Tst(): RptSHld LasOHYmd: End Sub

Sub RptSHld(A As Ymd)
Dim O$(): O = ZZOFxAy(A)
If AskOpnFxAy(O) Then Exit Sub
If Not IsMB52Loaded(A) Then Exit Sub
If IsCnlNoGit(A) Then Exit Sub
If IsCnlNoFc(YMzYmd(A)) Then Exit Sub
DltFfnAyIf O

UpdTbPHStkDays7_FldStkDays_andRemSc A ' Cannot put in Load_MB52, because, it is required to calculate the stock days to include the Git
Dim X As Excel.Application: Set X = NwXls
TmpPH5
UURpt X, CoYmd(86, A)
UURpt X, CoYmd(87, A)     ' The OFx will be created and Wb will be closed
OpnXFxy X, O
DrpCTT "#2Tot #A #B #Piv >MB52 #YpSku"
DrpCTblByPfxss "#StkHld $PH @Fc @StkHld @StkDays"
End Sub
'--------------------------------------
Private Sub UURpt__Tst()
TmpPH5
Dim X As Excel.Application: Set X = NwXls
UURpt X, CoYmd(86, LasOHYmd)
MaxiXls X
End Sub

Private Sub UURpt(X As Excel.Application, A As CoYmd)
Dim Wh$: Wh = OHYmdBexp(A.Ymd)
UUOupStkHld A   '@StkHld{7}
UUOupStkDays A  '@StkDays{7}
UUOupFc A       '@Fc{7}
UUGen X, A
Select Case 2
Case 1: RunCQ "Update Report x set DteGen=Now()" & Wh
Case 2
    With CurrentDb.OpenRecordset("Select DteGen from Report" & Wh)
        .Edit
        !DteGen = Now()
        .Update
    End With
End Select
DoCmd.SetWarnings True
End Sub
Private Sub UUOupStkHld__Tst()
Dim A As CoYmd: A = CoYmd(86, Ymd(19, 11, 29))
UUOupStkHld A
Exit Sub
#If False Then
Dim Wh$: Wh = CoOHYmdBexp(CoYmd(87, LasOHYmd))
RunCQ "SELECT Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(Val/1000) AS V" & _
" INTO [#A87]" & _
" FROM (OH x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & Wh

RunCQ "SELECT YpStk,x.Sku,Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(Val/1000) AS V" & _
" INTO [#A87]" & _
" FROM (OH x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & Wh & _
" GROUP BY YpStk,x.Sku;"

from GitDel
RunCQ "SELECT " & GitYpStk & ",x.Sku,Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(Amt/1000) AS V" & _
" INTO [#YpSku]" & _
" FROM (GitDet x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & Wh & _
" GROUP BY x.Sku;"

RunCQ "SELECT YpStk,x.Sku,Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(Val/1000) AS V" & _
" INTO [#A87]" & _
" FROM (OH x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & Wh & _
" GROUP BY YpStk,x.Sku;"

RunCQ "SELECT Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(Val/1000) AS V" & _
" INTO [#A86]" & _
" FROM (OH x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & CoOHYmdBexp(CoYmd(86, LasOHYmd))
Dim Sc86#, K86@, Sc87#, K87@
With CurrentDb.OpenRecordset("Select * from [#A86]")
    Sc86 = .Fields(0).Value
    K86 = .Fields(1).Value
End With
With CurrentDb.OpenRecordset("Select * from [#A87]")
    Sc87 = .Fields(0).Value
    K87 = .Fields(1).Value
End With
Debug.Print 86, Sc86, K86
Debug.Print 87, Sc87, K87
Debug.Print "Tot", Sc86 + Sc87, K87 + K87
#End If
End Sub
Private Sub UUOupStkHld(A As CoYmd)
'Aim: Create @StkHld{7} from Tbl-OH
'10+4Val = ScCsg HkdCsg ScDf HkdDf ScDp  HkdDp F1 ScGit HkdGit F2 ScTot HkdTot StkDays StkMth RemSC StkMthTar
'                    #StkHld{7}    AddPH7Atr
'                    {7Key}        {7Atr}
'Oup: @StkHldStm   = Stm         | Stream
'Oup: @StkHldBus   = Stm BusArea | Stream PHSBus BusArea PHBus
'Oup: @StkHldL1    = Stm PHL1    | Stream Srt1 PHL1 PHNam
'Oup: @StkHldL2    = Stm PHL2    | Stream Srt2 PHL2 PHNam PHBrd
'Oup: @StkHldL3    = Stm PHL3    | Stream Srt3 PHL3 PHNam PHBrd PHQGp
'Oup: @StkHldL4    = Stm PHL4    | Stream Srt4 PHL4 PHNam PHBrd PHQGp PHQly
'Oup: @StkHldSku   = Sku         | Stream Srt4 PHL4 PHNam PHBrd PHQGp PHQly Sku SkuDes

'     Stp       Oup
'     Beg       OH       = YY MM DD Co YpStk SKu BchNo | Btl Val
'  1   Sum       #YpSku   = YpStk Sku | Sc V                From OH
'  2   Piv       #Piv        = Sku   {8SC/Hkd}            Sku is not unique
'  3   2Tot      #2Tot       = Sku    |  {8SC/Hkd}  {2Tot}   Sku is unique
'  4   Rollup    #StkHld{7}  = {7Key} | {10Val}
'  5   Add4Kpi   @StkHld{7}  = Sku | {10Val} {4Days}
'  6   AddPH7Atr @StkHld{7}  = {7Atr} | {10Val}
'  7   AddF1F2   @StkHld{7}
'  8   ReSeq
'  9   DrpTmp
'Ref-For-AddAtr: $PH{5} PHLBus PHLStm = Sku | Btl/Sc ..

Dim WhYmd$: WhYmd = CoOHYmdBexp(A)
Dim WhCo$: WhCo = " where Co=" & A.Co
DoCmd.SetWarnings False
'== 1 Stp-Sum OH & Git

'PX "AC      =[@Btl] / [@[Btl/AC]]"
'PX "SC      =[@AC] * [@[Unit/AC]] / [@[Unit/SC]]"
'From OH
RunCQ "SELECT YpStk,x.Sku,Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(Val/1000) AS V" & _
" INTO [#YpSku]" & _
" FROM (OH x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & WhYmd & _
" GROUP BY YpStk,x.Sku;"

'From GitDet
RunCQ "Insert into [#YpSku] (YpStk,Sku,SC,V)" & _
" SELECT " & GitYpStk & " As YpStk ,x.Sku,Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(HKD/1000) AS V" & _
" FROM (GitDet x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & WhYmd & _
" GROUP BY x.Sku;"

If False Then
    Debug.Print A.Co, "SC/V",
    With CurrentDb.OpenRecordset("Select Sum(x.SC) as SC,Sum(x.V) as V from [#YpSku] x ")
        Debug.Print .Fields(0).Value; .Fields(1).Value
    End With
    Stop
End If

'-- 2 Stp-Piv =====================================================================================
DrpCT "#Piv"
RunCQ "Create Table [#Piv] (Sku Text(20)," & _
" ScCsg double, HkdCsg currency," & _
" ScDF  double, HkdDF  currency," & _
" ScDP  double, HkdDp  currency," & _
" ScGIT double, HkdGIT currency)"

RunCQ "INSERT INTO [#Piv] (Sku, HKDCsg, ScCsg) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*Consignment') group by Sku"
RunCQ "INSERT INTO [#Piv] (Sku, HKDDF , ScDF ) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*DutyFree')    group by Sku;"
RunCQ "INSERT INTO [#Piv] (Sku, HKDDP , ScDP ) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*DutyPaid')    group by Sku;"
RunCQ "INSERT INTO [#Piv] (Sku, HKDGit , ScGit) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*Git')    group by Sku;"

'-- 3 Stp-2Tot................................................................
RunCQ "SELECT Distinct x.Sku, " & _
" Sum(x.ScCsg) AS ScCsg, Sum(x.HKDCsg) AS HKDCsg," & _
" Sum(x.ScDF)  AS ScDF , Sum(x.HKDDF)  AS HKDDF  ," & _
" Sum(x.ScDP)  AS ScDP , Sum(x.HKDDP)  AS HKDDP  ," & _
" Sum(x.ScGIT) AS ScGIT, Sum(x.HKDGIT) AS HKDGIT" & _
" Into [#2Tot]" & _
" From [#Piv] x" & _
" Group by x.Sku"

RunCQ "Alter Table [#2Tot] add column ScTot Double, HKDTot Currency"
RunCQ "Update [#2Tot] set" & _
" ScTot=Nz(ScCsg,0)+Nz(ScDp,0)+Nz(ScDf,0)+Nz(ScGit,0)," & _
" HkdTot=Nz(HkdCsg,0)+Nz(HkdDp,0)+Nz(HkdDf,0)+Nz(HkdGit,0)"
If False Then
    With CurrentDb.OpenRecordset("Select Sum(ScTot),Sum(HkdTot) from [#2Tot]")
        Debug.Print "SC"; .Fields(0).Value
        Debug.Print "V "; .Fields(1).Value
    End With
    Stop
End If
'4 Stp-Rollup  ================================================================
'   Oup:#StkHld{7}
Const Sum$ = "Sum(x.ScCsg) as ScCsg,Sum(x.HkdCsg) as HkdCsg," & _
"Sum(x.ScDf) as ScDf,Sum(x.HkdDf) as HkdDf," & _
"Sum(x.ScDp) as ScDp,Sum(x.HkdDp) as HkdDp," & _
"Sum(x.ScGit) as ScGit,Sum(x.HkdGit) as HkdGit," & _
"Sum(x.ScTot) as ScTot,Sum(x.HkdTot) as HkdTot"

Add3PHRollupCol "#2Tot"
RunCQ "SELECT x.Sku,      " & Sum & " INTO [#StkHldSku] FROM [#2Tot] x Group by Sku"
RunCQ "SELECT Stm,BusArea," & Sum & " INTO [#StkHldBus] FROM [#2Tot] x Group by Stm,BusArea"
RunCQ "SELECT Stm,PHL4,   " & Sum & " INTO [#StkHldL4]  FROM [#2Tot] x Group by Stm,PHL4"
RunCQ "SELECT Stm,Left(PHL4,7) as PHL3," & Sum & " INTO [#StkHldL3] FROM [#StkHldL4] x Group by Stm,Left(PHL4,7)"
RunCQ "SELECT Stm,Left(PHL3,4) as PHL2," & Sum & " INTO [#StkHldL2] FROM [#StkHldL3] x Group by Stm,Left(PHL3,4)"
RunCQ "SELECT Stm,Left(PHL2,2) as PHL1," & Sum & " INTO [#StkHldL1] FROM [#StkHldL2] x Group by Stm,Left(PHL2,2)"
RunCQ "SELECT Stm,                     " & Sum & " INTO [#StkHldStm] FROM [#StkHldL1] x Group by Stm"
If False Then
    Stop
End If
'--5 Stp-Add4Kpi : StkDays/RemSC/StkMths/TarStkMths
'   ref: PHTarMths{7}
'   ref: PHStkDays{7}
Const AddCol$ = " add column StkDays integer, RemSC Double, StkMths Single, TarStkMths Single"
RunCQ "Alter Table [#StkHldSku]" & AddCol
RunCQ "Alter Table [#StkHldL4] " & AddCol
RunCQ "Alter Table [#StkHldL3] " & AddCol
RunCQ "Alter Table [#StkHldL2] " & AddCol
RunCQ "Alter Table [#StkHldL1] " & AddCol
RunCQ "Alter Table [#StkHldBus]" & AddCol
RunCQ "Alter Table [#StkHldStm]" & AddCol

'-- Sku ----------------------------------------
RunCQ "Select Sku,StkDays,RemSC into [#A] from PHStkDaysSku" & WhYmd
RunCQ "Select Sku,TarStkMths    into [#B] from PHTarMthsSKU" & WhCo
RunCQ "Update [#StkHldSku] x inner join [#A] a on a.Sku=x.Sku set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunCQ "Update [#StkHldSku] x inner join [#B] a on x.Sku=a.Sku set x.TarStkMths=a.TarStkMths"

'-- Bus ----------------------------------------
RunCQ "Select Stm,BusArea,StkDays,RemSC into [#A] from PHStkDaysBus" & WhYmd
RunCQ "Select Stm,BusArea,TarStkMths    into [#B] from PHTarMthsBus" & WhCo
RunCQ "Update [#StkHldBus] x inner join [#A] a on x.BusArea=a.BusArea and x.Stm=a.Stm set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunCQ "Update [#StkHldBus] x inner join [#B] a on x.BusArea=a.BusArea and x.Stm=a.Stm set x.TarStkMths=a.TarStkMths"

'-- PHL4 ----------------------------------------
RunCQ "Select Stm,PHL4,StkDays,RemSC into [#A] from PHStkDaysL4" & WhYmd
RunCQ "Select Stm,PHL4,TarStkMths    into [#B] from PHTarMthsL4" & WhCo
RunCQ "Update [#StkHldL4] x inner join [#A] a on x.PHL4=a.PHL4 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunCQ "Update [#StkHldL4] x inner join [#B] a on x.PHL4=a.PHL4 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- PHL3 ----------------------------------------
RunCQ "Select Stm,PHL3,StkDays,RemSC into [#A] from PHStkDaysL3" & WhYmd
RunCQ "Select Stm,PHL3,TarStkMths    into [#B] from PHTarMthsL3" & WhCo
RunCQ "Update [#StkHldL3] x inner join [#A] a on a.PHL3=x.PHL3 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunCQ "Update [#StkHldL3] x inner join [#B] a on x.PHL3=a.PHL3 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- PHL2 ----------------------------------------
RunCQ "Select Stm,PHL2,StkDays,RemSC into [#A] from PHStkDaysL2" & WhYmd
RunCQ "Select Stm,PHL2,TarStkMths    into [#B] from PHTarMthsL2" & WhCo
RunCQ "Update [#StkHldL2] x inner join [#A] a on a.PHL2=x.PHL2 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunCQ "Update [#StkHldL2] x inner join [#B] a on x.PHL2=a.PHL2 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- PHL1
RunCQ "Select Stm,PHL1,StkDays,RemSC into [#A] from PHStkDaysL1" & WhYmd
RunCQ "Select Stm,PHL1,TarStkMths    into [#B] from PHTarMthsL1" & WhCo
RunCQ "Update [#StkHldL1] x inner join [#A] a on a.PHL1=x.PHL1 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunCQ "Update [#StkHldL1] x inner join [#B] a on x.PHL1=a.PHL1 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- Stm
RunCQ "Select Stm,StkDays,RemSC into [#A] from PHStkDaysStm" & WhYmd
RunCQ "Select Stm,TarStkMths    into [#B] from PHTarMthsStm" & WhCo
RunCQ "Update [#StkHldStm] x inner join [#A] a on a.Stm=x.Stm set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunCQ "Update [#StkHldStm] x inner join [#B] a on x.Stm=a.Stm set x.TarStkMths=a.TarStkMths"

'-- 6 Stp-AddPH7Atr
'   Oup-@StkHld{7}
Const FmTblLik$ = "#StkHld?"
Const ToTblLik$ = "@StkHld?"
AddPH7Atr FmTblLik, ToTblLik

'-- 7 Stp-AddF1F2
Dim I: For Each I In PH7Ay: W2AddF1F2 "@StkHld" & I: Next

'-- 8 Stp-Reseq ===================================
SrtFldPH7 "@StkHld?", SH14ValFF

'-- 9 Stp-DrpTmp
If False Then
    DrpCTT "#A #B #YpSku #Piv #2Tot"
    DrpCTny PH7Tbny("#StkHld?")
End If
End Sub

Private Sub W2AddF1F2(T)
RunCQ "Alter Table [" & T & "] add column F1 Text(1),F2 Text(1)"
End Sub

Private Sub UUOupStkDays(A As CoYmd)
'       Lvl: #E*
'       Stm: Stm         Stream
'       Bus: Stm BusArea Stream PHSBus BusArea PHBus
'       L1 : Stm PHL1    Stream Srt1 PHNam
'       L2 : Stm PHL2    Stream Srt2 PHNam PHBrd
'       L3 : Stm PHL3    Stream Srt3 PHNam PHBrd PHQGp
'       L4 : Stm PHL4    Stream Srt4 PHNam PHBrd PHQGp PHQly
'       Sku: Sku         Stream Srt4 PHNam PHBrd PHQGp PHQly Sku SkuDes
'Inp: PHStkDaysStm = YY MM DD Co Stm         StkDays
'     PHStkDaysBus = YY MM DD Co Stm BusArea StkDays
'     PHStkDaysL1  = YY MM DD Co Stm PHL1    StkDays
'     PHStkDaysL2  = YY MM DD Co Stm PHL2    StkDays
'     PHStkDaysL3  = YY MM DD Co Stm PHL3    StkDays
'     PHStkDaysL4  = YY MM DD Co Stm PHL4    StkDays
'     PHStkDaysSku = YY MM DD Co Stm Sku     StkDays
'Stp  Oup   Des
'Key  #Key  = YY MM DD Co
'D    #D{7}
'E    #E{7}
'Tmp: #D: #Key YY MM DD Co
'Tmp#E
'     #EStm Stm     M01..15
'     #ESku Stm Sku M01..15
'     ..
'== Stp-Key = N YY MM DD, Where *YY *MM and roll 15 back from given @Ymd and *DD is the max with the *YY-&-*MM ==============
'Oup:#StkDays_NYmd = Co N YY MM DD                          ' DD is max(DD)
'Fm : PHStkDaysStm = YY MM DD Co Stm | ..

Dim Y As Byte: Y = A.Ymd.Y
Dim M As Byte: M = A.Ymd.M
RunCQ FmtStr("SELECT Top 15 Co,(YY-{0})*12+MM-{1} AS N, YY, MM, Max(x.DD) AS DDLng" & _
" INTO [#Key]" & _
" FROM PHStkDaysStm x" & _
" Where Co=" & A.Co & _
" Group BY Co,(YY-{0})*12+MM-{1},YY,MM" & _
" HAVING (YY-{0})*12+MM-{1} Between -14 And 0;", Y, M)
'== Adj DDLng to DD
RunCQ "Alter Table [#Key] add Column DD Byte"
RunCQ "Update [#Key] set DD = DDLng"
RunCQ "Alter Table [#Key] drop column DDLng"

'==Stp-D ===============================================================================================
'=#DSku
Const Sel$ = "Select x.N, x.YY, x.MM, x.DD, StkDays, RemSC, "
Const JnOn$ = " ON x.DD=a.DD AND x.MM=a.MM AND x.YY=a.YY and x.Co=a.Co;"
DoCmd.SetWarnings False
RunCQ Sel & "a.Sku        INTO [#DSku] FROM [#Key] x INNER JOIN PHStkDaysSku a" & JnOn
RunCQ Sel & "Stm          INTO [#DStm] FROM [#Key] x INNER JOIN PHStkDaysStm a" & JnOn
RunCQ Sel & "Stm, BusArea INTO [#DBus] FROM [#Key] x INNER JOIN PHStkDaysBus a" & JnOn
RunCQ Sel & "Stm, PHL1    INTO [#DL1]  FROM [#Key] x INNER JOIN PHStkDaysL1  a" & JnOn
RunCQ Sel & "Stm, PHL2    INTO [#DL2]  FROM [#Key] x INNER JOIN PHStkDaysL2  a" & JnOn
RunCQ Sel & "Stm, PHL3    INTO [#DL3]  FROM [#Key] x INNER JOIN PHStkDaysL3  a" & JnOn
RunCQ Sel & "Stm, PHL4    INTO [#DL4]  FROM [#Key] x INNER JOIN PHStkDaysL4  a" & JnOn

'== Stp-E================================================================================================
'#ESku
DrpCT "#ESku"
Const F15$ = "StkDays01 Integer, RemSC01 Double," & _
"StkDays02 Integer, RemSC02 Double," & _
"StkDays03 Integer, RemSC03 Double," & _
"StkDays04 Integer, RemSC04 Double," & _
"StkDays05 Integer, RemSC05 Double," & _
"StkDays06 Integer, RemSC06 Double," & _
"StkDays07 Integer, RemSC07 Double," & _
"StkDays08 Integer, RemSC08 Double," & _
"StkDays09 Integer, RemSC09 Double," & _
"StkDays10 Integer, RemSC10 Double," & _
"StkDays11 Integer, RemSC11 Double," & _
"StkDays12 Integer, RemSC12 Double," & _
"StkDays13 Integer, RemSC13 Double," & _
"StkDays14 Integer, RemSC14 Double," & _
"StkDays15 Integer, RemSC15 Double)"

RunCQ "Create Table [#ESku] (Sku Text(20)," & F15
RunCQ "INSERT INTO [#ESku] SELECT Distinct Sku From [#DSku]"
Dim N%: For N = 0 To -14 Step -1
Dim NStr$: NStr = Format(1 - N, "00")
RunCQ "UPDATE [#ESku] x INNER JOIN [#DSku] a ON x.Sku=a.Sku SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EStm ..............................................................................................
DrpCT "#EStm"
RunCQ "Create Table [#EStm] (Stm Text(1)," & F15
RunCQ "INSERT INTO [#EStm] SELECT Distinct Stm From [#DStm]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunCQ "UPDATE [#EStm] x INNER JOIN [#DStm] a ON x.Stm=a.Stm SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EBus ..............................................................................................
DrpCT "#EBus"
RunCQ "Create Table [#EBus] (Stm Text(1),BusArea Text(4)," & F15
RunCQ "INSERT INTO [#EBus] SELECT Distinct Stm,BusArea From [#DBus]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunCQ "UPDATE [#EBus] x INNER JOIN [#Dbus] a ON x.Stm=a.Stm and x.BusArea=a.BusArea SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL1 ..............................................................................................
DrpCT "#EL1"
RunCQ "Create Table [#EL1] (Stm Text(1),PHL1 Text(2)," & F15
RunCQ "INSERT INTO [#EL1] SELECT Distinct Stm,PHL1 From [#DL1]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunCQ "UPDATE [#EL1] x INNER JOIN [#DL1] a ON x.Stm=a.Stm and x.PHL1=a.PHL1 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL2 ..............................................................................................
DrpCT "#EL2"
RunCQ "Create Table [#EL2] (Stm Text(1),PHL2 Text(4)," & F15
RunCQ "INSERT INTO [#EL2] SELECT Distinct Stm,PHL2 From [#DL2]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunCQ "UPDATE [#EL2] x INNER JOIN [#DL2] a ON x.Stm=a.Stm and x.PHL2=a.PHL2 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL3 ..............................................................................................
DrpCT "#EL3"
RunCQ "Create Table [#EL3] (Stm Text(1),PHL3 Text(7)," & F15
RunCQ "INSERT INTO [#EL3] SELECT Distinct Stm,PHL3 From [#DL3]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunCQ "UPDATE [#EL3] x INNER JOIN [#DL3] a ON x.Stm=a.Stm and x.PHL3=a.PHL3 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL4 ..............................................................................................
DrpCT "#EL4"
RunCQ "Create Table [#EL4] (Stm Text(1),PHL4 Text(10)," & F15
RunCQ "INSERT INTO [#EL4] SELECT Distinct Stm,PHL4 From [#DL4]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunCQ "UPDATE [#EL4] x INNER JOIN [#DL4] a ON x.Stm=a.Stm and x.PHL4=a.PHL4 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'==Stp-AddSC====================================================================================================================/
'-- #E{7} Add fields SC to each
'   From @StkHld{7}
RunCQ "Alter Table [#EStm] add column SC double"
RunCQ "Alter Table [#ESku] add column SC double"
RunCQ "Alter Table [#EBus] add column SC double"
RunCQ "Alter Table [#EL1] add column SC double"
RunCQ "Alter Table [#EL2] add column SC double"
RunCQ "Alter Table [#EL3] add column SC double"
RunCQ "Alter Table [#EL4] add column SC double"

'== Stp-OHSC{7} from @StkHld{7}
'-- Create #OHSC{7} from @StkHld{7}
RunCQ "Select Sku           ,ScTot As SC into [#OHSCSku] from [@StkHldSku]"
RunCQ "Select Stm        ,ScTot As SC into [#OHSCStm] from [@StkHldStm] x inner join PHLStm a on a.Stream=x.Stream"
RunCQ "Select Stm,BusArea,ScTot As SC into [#OHSCBus] from [@StkHldBus] x inner join PHLStm a on a.Stream=x.Stream"
RunCQ "Select Stm,PHL1   ,ScTot As SC into [#OHSCL1]  from [@StkHldL1] x inner join PHLStm a on a.Stream=x.Stream"
RunCQ "Select Stm,PHL2   ,ScTot As SC into [#OHSCL2]  from [@StkHldL2] x inner join PHLStm a on a.Stream=x.Stream"
RunCQ "Select Stm,PHL3   ,ScTot As SC into [#OHSCL3]  from [@StkHldL3] x inner join PHLStm a on a.Stream=x.Stream"
RunCQ "Select Stm,PHL4   ,ScTot As SC into [#OHSCL4]  from [@StkHldL4] x inner join PHLStm a on a.Stream=x.Stream"
RunCQ "Update [#ESku] x inner join [#OHSCSku] a on a.Sku=x.Sku                         set x.SC=a.SC"
RunCQ "Update [#EStm] x inner join [#OHSCStm] a on a.Stm=x.Stm                         set x.SC=a.SC"
RunCQ "Update [#EBus] x inner join [#OHSCBus] a on a.Stm=x.Stm and a.BusArea=x.BusArea set x.SC=a.SC"
RunCQ "Update [#EL1]  x inner join [#OHSCL1]  a on a.Stm=x.Stm and a.PHL1   =x.PHL1    set x.SC=a.SC"
RunCQ "Update [#EL2]  x inner join [#OHSCL2]  a on a.Stm=x.Stm and a.PHL2   =x.PHL2    set x.SC=a.SC"
RunCQ "Update [#EL3]  x inner join [#OHSCL3]  a on a.Stm=x.Stm and a.PHL3   =x.PHL3    set x.SC=a.SC"
RunCQ "Update [#EL4]  x inner join [#OHSCL4]  a on a.Stm=x.Stm and a.PHL4   =x.PHL4    set x.SC=a.SC"

'== Stp-InsWithOHNoStkDays
'   #E{7}->SC insert records where has OH, but no StkDays
RunCQ "Insert into [#ESku] (Sku        ,SC) select x.Sku          ,x.SC from [#OHScSku] x left join [#ESku] a on x.Sku=a.sku                   where a.Sku is null"
RunCQ "Insert into [#EL4]  (Stm,PHL4   ,SC) select x.Stm,x.PHL4   ,x.SC from [#OHScL4]  x left join [#EL4]  a on x.PHL4=a.PHL4 and x.Stm=a.Stm where a.Stm is null"
RunCQ "Insert into [#EL3]  (Stm,PHL3   ,SC) select x.Stm,x.PHL3   ,x.SC from [#OHScL3]  x left join [#EL3]  a on x.PHL3=a.PHL3 and x.Stm=a.Stm where a.Stm is null"
RunCQ "Insert into [#EL2]  (Stm,PHL2   ,SC) select x.Stm,x.PHL2   ,x.SC from [#OHScL2]  x left join [#EL2]  a on x.PHL2=a.PHL2 and x.Stm=a.Stm where a.Stm is null"
RunCQ "Insert into [#EL1]  (Stm,PHL1   ,SC) select x.Stm,x.PHL1   ,x.SC from [#OHScL1]  x left join [#EL1]  a on x.PHL1=a.PHL1 and x.Stm=a.Stm where a.Stm is null"
RunCQ "Insert into [#EBus] (Stm,BusArea,SC) select x.Stm,x.BusArea,x.SC from [#OHScBus] x left join [#EBus] a on x.BusArea=a.BusArea and x.Stm=a.Stm where a.Stm is null"

'== Stp-AddAtr =============================================================================================
'-- Add Attribute Fields
AddPH7Atr "#E?", "@StkDays?"

'== Stp-Reseq ======================================================================================================
Const RstFlds$ = "SC" & _
" StkDays01 RemSC01" & _
" StkDays02 RemSC02" & _
" StkDays03 RemSC03" & _
" StkDays04 RemSC04" & _
" StkDays05 RemSC05" & _
" StkDays06 RemSC06" & _
" StkDays07 RemSC07" & _
" StkDays08 RemSC08" & _
" StkDays09 RemSC09" & _
" StkDays10 RemSC10" & _
" StkDays11 RemSC11" & _
" StkDays12 RemSC12" & _
" StkDays13 RemSC13" & _
" StkDays14 RemSC14" & _
" StkDays15 RemSC15"
SrtFldPH7 "@StkDays?", RstFlds

'== Stp-DrpTmp ======================================================================================================
Dim I
For Each I In Split("Sku Stm Bus L1 L2 L3 L4")
    RunCQ "Drop Table [#D" & I & "]"
    RunCQ "Drop Table [#E" & I & "]"
    RunCQ "Drop Table [#OHSc" & I & "]"
Next
RunCQ "Drop Table [#Key]"
End Sub

Private Sub UUOupFc(A As CoYmd)
'Aim: create @Fc{7} From FcSku
'Oup: @Fc{7}   {7LvlKey} | SC | M01..15

'## Stp       Oup      What
' 1 !@TmpFc     $Fc{7}   By calling TmpFc_ByCoYM
' 2 !@TmpScOH   $ScOH{7} By Call TmpScOH7_ByCoYmd
' 3 !@AddScCol  $Fc{7}   add a column SC
' 4 !@InsOHNoSc $Fc{7}   add records from $ScOH for those with OH, but not Fc
'   !@AddSdRemCol $Fc{7} add two columns StkDays and RemSC to each lvl from PHStkDays{7}
' 5 !@AddAtr    @Fc{7}
' 6 !@DrpCoCol
'   !@ReSeq     @Fc{7}
' 7 !@DrpTmp    $Fc{7} $ScOH{7}
'== 1 !@TmpFc
TmpFc_ByCoYM CoYMzCoYmd(A)

'== 2 !@TmpScOH
TmpScOH7_ByCoYmd A

'== 3 !@AddScCol
RunCQ "Alter Table [$FcSku] Add Column SC Double"
RunCQ "Alter Table [$FcL4] Add Column SC Double"
RunCQ "Alter Table [$FcL3] Add Column SC Double"
RunCQ "Alter Table [$FcL2] Add Column SC Double"
RunCQ "Alter Table [$FcL1] Add Column SC Double"
RunCQ "Alter Table [$FcBus] Add Column SC Double"
RunCQ "Alter Table [$FcStm] Add Column SC Double"

'-- Update $Fc{7}->SC
RunCQ "Update [$FcSku] x inner join [$ScOHSku] a on x.Co=a.Co and x.Sku=a.Sku         set x.SC = a.SC"
RunCQ "Update [$FcL4]  x inner join [$ScOHL4]  a on x.Co=a.Co and x.PHL4=a.PHL4       set x.SC = a.SC"
RunCQ "Update [$FcL3]  x inner join [$ScOHL3]  a on x.Co=a.Co and x.PHL3=a.PHL3       set x.SC = a.SC"
RunCQ "Update [$FcL2]  x inner join [$ScOHL2]  a on x.Co=a.Co and x.PHL2=a.PHL2       set x.SC = a.SC"
RunCQ "Update [$FcL1]  x inner join [$ScOHL1]  a on x.Co=a.Co and x.PHL1=a.PHL1       set x.SC = a.SC"
RunCQ "Update [$FcBus] x inner join [$ScOHBus] a on x.Co=a.Co and x.BusArea=a.BusArea set x.SC = a.SC"
RunCQ "Update [$FcStm] x inner join [$ScOHStm] a on x.Co=a.Co and x.Stm=a.Stm         set x.SC = a.SC"
'== 4 !@InsOHNoFc
RunCQ "Insert into [$FcSku] (Co,Sku        ,SC) select x.Co,x.Sku          ,x.SC from [$ScOHSku] x left join [$FcSku] a on x.Co=a.Co and x.Sku=a.sku                   where a.Co is null"
RunCQ "Insert into [$FcL4]  (Co,Stm,PHL4   ,SC) select x.Co,x.Stm,x.PHL4   ,x.SC from [$ScOHL4]  x left join [$FcL4]  a on x.Co=a.Co and x.PHL4=a.PHL4 and x.Stm=a.Stm where a.Co is null"
RunCQ "Insert into [$FcL3]  (Co,Stm,PHL3   ,SC) select x.Co,x.Stm,x.PHL3   ,x.SC from [$ScOHL3]  x left join [$FcL3]  a on x.Co=a.Co and x.PHL3=a.PHL3 and x.Stm=a.Stm where a.Co is null"
RunCQ "Insert into [$FcL2]  (Co,Stm,PHL2   ,SC) select x.Co,x.Stm,x.PHL2   ,x.SC from [$ScOHL2]  x left join [$FcL2]  a on x.Co=a.Co and x.PHL2=a.PHL2 and x.Stm=a.Stm where a.Co is null"
RunCQ "Insert into [$FcL1]  (Co,Stm,PHL1   ,SC) select x.Co,x.Stm,x.PHL1   ,x.SC from [$ScOHL1]  x left join [$FcL1]  a on x.Co=a.Co and x.PHL1=a.PHL1 and x.Stm=a.Stm where a.Co is null"
RunCQ "Insert into [$FcBus] (Co,Stm,BusArea,SC) select x.Co,x.Stm,x.BusArea,x.SC from [$ScOHBus] x left join [$FcBus] a on x.Co=a.Co and x.BusArea=a.BusArea and x.Stm=a.Stm where a.Co is null"

'==   !@Add_StkDays_RemSC_Col
Dim Wh$: Wh = CoOHYmdBexp(A)
Dim Jn$(): Jn = PH7Jn
Dim K$(): K = PH7Key
Dim I, J%: For Each I In PH7Ay
    Add_StkDays_RemSC_Col I, K(J), Jn(J), Wh
    J = J + 1
Next

'== 5 !@AddAtr
AddPH7Atr "$Fc?", "@Fc?"

'== 6 !@DrpCoCol
For Each I In PH7Ay
    RunCQ "alter Table [@Fc" & I & "] drop column Co"
Next
RunCQ "Alter Table [@FcSku] drop column Stm"
'== 7 !@Reseq
SrtFldPH7 "@Fc?", "SC StkDays RemSC" & _
" M01 M02 M03" & _
" M04 M05 M06" & _
" M07 M08 M09" & _
" M10 M11 M12" & _
" M13 M14 M15"

'== !@DrpTmp
For Each I In PH7Ay
    RunCQ "Drop Table [$Fc" & I & "]"
    RunCQ "Drop Table [$ScOH" & I & "]"
Next
End Sub
Private Sub UUGen__Tst()
Dim X As Excel.Application: Set X = NwXls
UUGen X, CoYmd(86, LasOHYmd)
UUGen X, CoYmd(87, LasOHYmd)
X.WindowState = xlMaximized
X.Windows.Arrange xlArrangeStyleVertical
Done
End Sub
Private Sub UUGen(X As Excel.Application, A As CoYmd)
Dim OFx$: OFx = ShOFxzCoYmd(A)
CpyFfn ShTp, ShOFxzCoYmd(A)
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(OFx)
RfhWb Wb, CFb
W3FmtShRpt Wb, A
Wb.Save
Wb.Close
End Sub

Private Sub Add_StkDays_RemSC_Col(PHItm, PHKey$, PHJn$, Wh$)
Dim I$: I = PHItm
RunCQ "Alter Table [$Fc" & I & "] add Column StkDays Integer, RemSC double"
RunCQ "Select " & PHKey & ",StkDays,RemSC Into [#A] from PHStkDays" & I & " " & Wh
RunCQ "Update [$Fc" & I & "] x inner Join [#A] a on " & PHJn & " Set x.StkDays=a.StkDays,x.RemSC=a.RemSC"
End Sub

Function W3FmtShRpt(Wb As Workbook, A As CoYmd)
MinvWb Wb ' This is need because, Merge & Unmerge will break under MiniState
FmtSdDteTit Wb, YMzYmd(A.Ymd)
FmtFcDteTit Wb, YMzYmd(A.Ymd)
MiniWb Wb
W3FmtA1 Wb, A
End Function

Private Sub W3FmtA1__Tst()
Dim Wb As Workbook: Set Wb = ShTpWb
W3FmtA1 Wb, CoYmd(86, Ymd(19, 12, 3))
Wb.Application.WindowState = xlMaximized
End Sub
Private Sub W3FmtA1(Wb As Workbook, A As CoYmd)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If IsShRpt(Ws.Name) Or Ws.Name = "Index" Then
        W1SetA1 Ws.Range("A1"), A
    End If
Next
Set Ws = Wb.Sheets("Index")
Ws.Range("C2").Value = Now
End Sub
Private Function W1A1Sfx$(A As CoYmd)
Dim Co$: Co = CoNm(A.Co)
W1A1Sfx = " As At " & YYmdStr(A.Ymd) & " (" & Co & ")"
End Function
Private Sub W1SetA1(A1 As Range, A As CoYmd)
Dim OldA1$: OldA1 = A1.Value
Dim NwA1$: NwA1 = BefOrAll(OldA1, " As At") & W1A1Sfx(A)
A1.Value = NwA1
End Sub

Private Function ZZSampCoYmd() As CoYmd: ZZSampCoYmd = CoYmd(86, Ymd(19, 11, 29)): End Function
Private Function ZZOFxAy(A As Ymd) As String()
Dim O$(1)
O(0) = ShOFx(86, A)
O(1) = ShOFx(87, A)
ZZOFxAy = O
End Function


