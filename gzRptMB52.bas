Attribute VB_Name = "gzRptMB52"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRptMB52."
Public Const SapDtaSS$ = "Litre/Btl Btl/AC Unit/AC Unit/SC"
'Sub T_TmpOH(): TmpOH__Tst: End Sub
'Sub T_VVOupOH(): VVOupOH__Tst: End Sub
'Sub T_RGen():  RGen__Tst:  End Sub

Private Sub RptMB52__Tst(): RptMB52 LasOHYmd: End Sub
Sub RptLasMB52(): RptMB52 LasOHYmd: End Sub
Sub RptMB52(A As Ymd) 'Create !MB52OFx from Tb-[]
Dim OFx$: OFx = MB52OFx(A)
If OpnFxIfExist(OFx) Then Exit Sub
DltFfnIf OFx
VVOupZHT0Rat
VVOupPacRat A
VVOupBchRatHK A '@BchRatHK: Using Sku BchNo BchRat
VVOupBchRatMO A '@BchRatMO: Using Sku BchNo BchRat
VVOupOH A
VVGenOFx OFx
VVUpdRpt A
DrpCTT "#OH @BchRatHK @BchRatMO @OH @PacRat @PacRatD @ZHT0Rat >MB52"
End Sub

Private Sub VVGenOFx(Fx$)
Dim Wb As Workbook: Set Wb = GenOupFx(Fx, MB52Tp, CFb)
SetLoFml FstLo(Wb.Sheets("Data")), MB52DtaWsFml
End Sub
Function MB52DtaWsFml() As String()
Dim O$()
'Unit
PushS O, "Btl/Unit =[@[Btl/AC]]/[@[Unit/AC]]"

'Siz
PushS O, "ml/Btl   =[@[Litre/Btl]]*1000"
PushS O, "Btl/AC'  =[@[Btl/AC]]"
PushS O, "Litre/SC =[@[Litre/Btl]] * [@[Btl/AC]] / [@[Unit/AC]] * [@[Unit/SC]]"

'OH
PushS O, "AC      =[@Btl] / [@[Btl/AC]]"
PushS O, "SC      =[@AC] * [@[Unit/AC]] / [@[Unit/SC]]"

'Pri
PushS O, "BtlUPr  =[@Val]/[Btl]"
PushS O, "AcUPr   =[@Val]/[@AC]"
PushS O, "ScUPr   =[@Val]/[@SC]"

'TaxItm
'Taxed
PushS O, "TaxItm = IF(OR(NOT(ISBLANK([@BchRat])),NOT(ISBLANK([@ZHT0Rat]))),""Y"","""")"
PushS O, "Taxed = IF(AND([@NoTax]<>""Y"",OR([@TaxItm]=""Y"",[@[3p]]=""Y""),[@TaxLoc]=""Y""),""Y"","""")"

'Amt
PushS O, "BchAmt  =IF([@Taxed]=""Y"",[@BchRat]*[@AC],0)"
PushS O, "ZHT0Amt =IF([@Taxed]=""Y"",[@ZHT0Rat]*[@AC],0)"

'Dif
PushS O, "RatDif  =Round([@BchRat]-[@ZHT0Rat],0)"
PushS O, "AmtDif  =Round(IF([@Taxed]=""Y"",[@BchAmt]-[@ZHT0Amt],0),0)"
MB52DtaWsFml = O
End Function

Private Sub VVUpdRpt(A As Ymd)
RunCQ "Update Report Set DteMB52Gen=#" & Now & "# where " & OHYmdBexp(A)
End Sub

Private Sub VVOupOH__Tst()
Dim A As Ymd: A = Ymd(19, 11, 29)
ClsAllCTbl
VVOupZHT0Rat
VVOupBchRatHK A '@BchRatHK: Using Sku BchNo BchRat
VVOupBchRatMO A '@BchRatMO: Using Sku BchNo BchRat
VVOupOH A
DoCmd.OpenTable "@OH"
End Sub
Private Sub VVOupOH(A As Ymd)
'Aim: Create #OH from OH & GITDet & Add Atr-fields
'Inp: OH     = YY MM DD Co Sku YpStk BchNo | Btl Val
'Inp: GITDet = YY MM DD Co Sku             | Btl HKD
'Ref: Sku    : ZHT0-Rate  = Sku->TaxRate    (/AC)
'     SkuB   : Sku BchNo | DutyRateB        ' DutyRateB is /Btl rate
'     YpStk  = YpStk | NmYpStk Co SLoc YpCls IsTaxLoc
' ## Stp   Oup
'Inp: #OH =
'Oup: @OH =
'Why: Just Cpy #OH by order and Rename fields
DoCmd.SetWarnings False
StsQry "@OH":
Dim Git%: Git = GitYpStk
Dim Wh$: Wh = WhOHYmd(A)

'Crt #OH  ! OH
'Ins #OH  ! GIT
RunCQ "Select" & _
                        " YY,MM,DD,Co,YpStk,SLoc,Sku,BchNo,Btl,Val Into [#OH] From OH" & Wh
RunCQ "Insert into [#OH] (YY,MM,DD,Co,YpStk,     Sku,Btl,Val) Select YY,MM,DD,Co," & Git & ",Sku,Btl,HKD from GITDet" & Wh

'AddCol #OH ! ..
RunCQ "Alter Table [#OH] add column " & _
"NmYpStk Text(50),TaxLoc Text(1),[3p] Text(1), NoTax Text(1)," & _
"SkuDes Text(255),PH Text(20), CdTopaz Text(20), Topaz Long," & _
"[Litre/Btl] Double," & _
"[Btl/AC] Integer  ," & _
"[Unit/AC] Integer ," & _
"[Unit/SC] Double  ," & _
"StkUnit Text(3)   , " & _
"BusArea Text(40), PHL1 Text(2),PHL2 Text(4),PHL3 Text(7),PHL4 Text(10)," & _
"Stream Text(6),PHBus Text(99), PHNam Text(99),PHBrd Text(99),PHQGp Text(99),PHQly Text(99)," & _
"PHSStm Integer, PHSBus Integer, PHSrt1 Text(2), PHSrt2 Text(4), PHSrt3 Text(6), PHSrt4 Text(8)," & _
"BchRat Double, BchRatTy Text(4), ZHT0Rat Double"
Const UpdJn$ = "Update [#OH] x inner join "
Const UpdSet$ = "Update [#OH] Set "

'UpdFld #OH->YpStk TaxLoc
'UpdFld #OH->3p
'UpdFld #OH->NoTax
RunCQ UpdJn & "YpStk a on a.YpStk=x.YpStk        set x.NmYpStk = a.NmYpStk, x.TaxLoc=IIf(a.IsTaxLoc,'Y','')"
RunCQ UpdJn & "SkuTaxBy3rdParty a on a.Sku=x.Sku set x.3p='Y'"
RunCQ UpdJn & "SkuNoLongerTax a on a.Sku=x.Sku   set x.NoTax='Y'"

'UpdFld #OH->PH SkuDes Topaz ...
RunCQ UpdJn & "Sku   a on a.Sku=x.Sku" & _
" set " & _
"x.PH=a.ProdHierarchy, x.SkuDes=a.SkuDes,x.Topaz=a.Topaz," & _
"x.[Litre/Btl]=a.[Litre/Btl]," & _
"x.[Btl/AC]   =a.[Btl/AC]   ," & _
"x.[Unit/AC]  =a.[Unit/AC]  ," & _
"x.[Unit/SC]  =a.[Unit/SC]  ," & _
"x.StkUnit    =a.StkUnit    ," & _
"x.BusArea    =a.BusArea"

RunCQ UpdJn & "Topaz a on a.Topaz=x.Topaz set x.CdTopaz=a.CdTopaz"
RunCQ UpdSet & "Stream=IIf(Left(CdTopaz,3)='UDV','Diageo','MH')"
RunCQ "Alter Table [#OH] Drop Column Topaz"

'Stp-PH ===========================================================================================
'Set PH1..4                     #OH
'Set PHNam..Qly PHSrt1..4       #PH?
'Set PHBus PHSBus               PHLBus
'Set PHSStm                     PHLStm
RunCQ UpdSet & "PHL1 = Left(PH,2), PHL2 = Left(PH,4), PHL3 = Left(PH,7), PHL4 = Left(PH,10)"

RunCQ "Select Left(PH,2) as PHL1,  Des as PHNam, Srt As PHSrt1 Into [#PHL1] from ProdHierarchy where Lvl=1"
RunCQ "Select Left(PH,4) as PHL2,  Des as PHBrd, Srt As PHSrt2 Into [#PHL2] from ProdHierarchy where Lvl=2"
RunCQ "Select Left(PH,7) as PHL3,  Des as PHQGp, Srt As PHSrt3 Into [#PHL3] from ProdHierarchy where Lvl=3"
RunCQ "Select Left(PH,10) as PHL4, Des as PHQly, Srt As PHSrt4 Into [#PHL4] from ProdHierarchy where Lvl=4"

RunCQ UpdJn & "[#PHL1] a on x.PHL1=a.PHL1 set x.PHNam=a.PHNam, x.PHSrt1=a.PHSrt1"
RunCQ UpdJn & "[#PHL2] a on x.PHL2=a.PHL2 set x.PHBrd=a.PHBrd, x.PHSrt2=a.PHSrt2"
RunCQ UpdJn & "[#PHL3] a on x.PHL3=a.PHL3 set x.PHQGp=a.PHQGp, x.PHSrt3=a.PHSrt3"
RunCQ UpdJn & "[#PHL4] a on x.PHL4=a.PHL4 set x.PHQly=a.PHQly, x.PHSrt4=a.PHSrt4"

DrpCTT ExpandPfxNN("#PHL", 1, 4)

RunCQ UpdJn & "PHLStm a on a.Stream=x.Stream   Set x.PHSStm=a.PHSStm"
RunCQ UpdJn & "PHLBus a on a.BusArea=x.BusArea Set x.PHBus=a.PHBus, x.PHSBus=a.PHSBus"

'AddCol #OH->BchRat BchRatTy Co86
'AddCol #OH->BchRat BchRatTy Co87
RunCQ UpdJn & "[@BchRatHK] a on a.Sku=x.Sku and a.BchNo=x.BchNo Set x.BchRat=a.BchRat,x.BchRatTy=a.BchRatTy Where x.Co=86 and TaxLoc='Y'"
RunCQ UpdJn & "[@BchRatMO] a on a.Sku=x.Sku and a.BchNo=x.BchNo Set x.BchRat=a.BchRat,x.BchRatTy='*MO'      Where x.Co=87 and TaxLoc='Y'"
RunCQ UpdJn & "[@ZHT0Rat] a on x.Sku=a.Sku AND x.Co=a.Co set x.ZHT0Rat=a.ZHT0Rat where TaxLoc='Y'"
'AddFld #OH-> All Fml Fld
Dim F: For Each F In T1Ay(TRstAy(MB52DtaWsFml))
    RunCQ "Alter Table [#OH] Add Column [" & F & "] Byte"
Next
NewRseqTbl "OH", W1Fny, "Co,PHSStm,PHSrt4,[Litre/Btl],Sku"
End Sub

Private Function W1Fny() As String()
Dim O$()
W1Push O, "Co YY MM DD"
W1Push O, "Stream PHBus BusArea NmYpStk YpStk"
W1Push O, "SkuDes Sku"

W1Push O, "Litre/Btl Btl/AC Unit/SC Unit/AC StkUnit"
W1Push O, "Btl/Unit"
W1Push O, "ml/Btl Btl/AC' Litre/SC"
W1Push O, "Val Btl AC SC"

W1Push O, "BtlUPr"
W1Push O, "AcUPr"
W1Push O, "ScUPr"

W1Push O, "SLoc TaxLoc TaxItm NoTax 3p Taxed"
W1Push O, "BchRat ZHT0Rat"
W1Push O, "BchAmt ZHT0Amt"
W1Push O, "RatDif AmtDif"

W1Push O, "BchNo BchRatTy"

W1Push O, "CdTopaz PH"
W1Push O, "PHNam PHBrd PHQGp PHQly"
W1Push O, "PHSStm PHSBus PHSrt1 PHSrt2 PHSrt3 PHSrt4"
W1Push O, "PHL1 PHL2 PHL3 PHL4"
W1Fny = O
End Function

Private Sub W1Push(O$(), SS$)
PushIAy O, SyzSS(SS)
End Sub


'------------------------------------------------------------------------------------------------------------------------ VVOupOH
Private Sub MB52DtaWsFml__Tst()
Dim Wb As Workbook: Set Wb = WbzFx(MB52Tp)
SetLoFml FstLo(Wb.Sheets("Data")), MB52DtaWsFml
MaxiWb Wb
End Sub

Private Sub SetBchRatMoFml(L As ListObject)
SetLcFmlln L, "Litre =[@Btl] * [@Litre/Btl]"
SetLcFmlln L, "LitreHKD =[@Litre] * [@[MOP/Litre]] / [@[HKD/MOP]]"
SetLcFmlln L, "10 A=[@[XXX/Btl]] * [@Btl] * 0.1"
SetLcFmlln L, "10 B=[@Val] * 0.1"
SetLcFmlln L, "10%HKD=IF(ISBLANK([@[XXX/Btl]]),[@[10%B]],[@[10%A]])"
SetLcFmlln L, "HKD=[@LitreHKD] + [@[10%HKD]]"
SetLcFmlln L, "HKD/AC=[@[Btl/AC]]"
End Sub

Private Sub VVOupBchRatHK__Tst()
Dim A As Ymd: A = Ymd(19, 11, 29)
ClsAllCTbl
VVOupPacRat A
VVOupBchRatHK A
DoCmd.OpenTable "@BchRatHK"
End Sub
Private Sub VVOupBchRatHK(A As Ymd)
'@WhYYMMDD ! Where Ymd
'Crt @BchRatHK : Co BchRatTy BchRat Sku Des BchNo BchPermitD BchPermit BchPermitDate LasPermitD LasPermit LasPermitDate LasBchNo FmSkuCnt
'   From  [SkuTaxBy3rdParty]  Sku RateU
'   From  [SkuNoLongerTax]    Sku
'   From  [$OHBch]  Sku BchNo      ! From OH where Co=86 & Btl>0
'   From  [@PacRat] NewSku Rate                                      ! Rate is /Btl
'   From  [$BchRat] Sku BchNo Rate PermitD Permit PermitDate           ! Rate is /Btl
'   From  [$LasRat] Sku       Rate PermitD Permit PermitDate BchNo  ! Rate is /Btl

StsQry "@BchRatHK"
DoCmd.SetWarnings False
W2TmpOHBch A
W2TmpBchRat A
W2TmpLasRat A

'Crt @BchRatHK : empty
DrpCT "@BchRatHK"
RunCQ "Create Table [@BchRatHK] (Co Byte, BchRatTy Text(6), BchRat Currency, Sku Text(15), Des Text(255),BchNo Text(10)," & _
" BchPermitD Long, BchPermit Long, BchPermitDate Date," & _
" LasPermitD Long, LasPermit Long, LasPermitDate date, LasBchNo Text(10)," & _
" FmSkuCnt Byte, BtlRat Currency, [Btl/Ac] byte)"

'Ins @BchRatHK : for all OH
RunCQ "Insert into [@BchRatHK] (Co,Sku,BchNo) select 86,Sku,BchNo from [$OHBch]"

'Crt #3p       : Sku RateU BltRat [Blt/Ac]
RunCQ "Select x.Sku,RateU,[Btl/Ac],CDbl(0) aS BtlRat into [#3p] from [SkuTaxBy3rdParty] x inner join Sku a on a.Sku=x.Sku"
RunCQ "Update [#3p] set BtlRat =RateU/[Btl/Ac]"  ' RateU is in /Ac
'Upd @BchRatHK : for *3p
'Upd @BchRatHK : for *NoTax
'Upd @BchRatHK : for *Bch
'Upd @BchRatHK : for *Las
'Upd @BchRatHK : for *Pac
RunCQ "Update [@BchRatHK] x inner join [#3p]            a on x.Sku=a.Sku              set BchRatTy='*3p' ,x.BtlRat=a.BtlRat"
RunCQ "Update [@BchRatHK] x inner join [SkuNoLongerTax] a on x.Sku=a.Sku              set BchRatTy='*NoTax'"
RunCQ "Update [@BchRatHK] x inner join [$BchRat] a on x.Sku=a.Sku and x.BchNo=a.BchNo set BchRatTy='*Bch',x.BtlRat=a.Rate,BchPermitD=PermitD,BchPermit=Permit,BchPermitDate=PermitDate"
RunCQ "Update [@BchRatHK] x inner join [$LasRat] a on x.Sku=a.Sku                     set BchRatTy='*Las',x.BtlRat=a.Rate,LasPermitD=PermitD,LasPermit=Permit,LasPermitDate=PermitDate,LasBchNo=a.BchNo where BchRatTy is null"
RunCQ "Update [@BchRatHK] x inner join [@PacRat] a on x.Sku=a.NewSku                  set BchRatTy='*Pac',x.BtlRat=a.Rate,x.FmSkuCnt=a.FmSkuCnt"

'Upd @BchRatHK : ->[Btl/Ac], Des
'Dlt @BchRatHK : for no BchRatTy
'Upd @BchRatHK : ->[BchRat]
RunCQ "Update [@BchRatHK] x inner join Sku a on x.Sku=a.Sku set x.[Btl/AC]=a.[Btl/Ac],x.Des=a.SkuDes"
RunCQ "Delete * from [@BchRatHK] where BchRatTy is null or Nz([Btl/Ac],0)=0"
RunCQ "Update [@BchRatHK] set BchRat=BtlRat*[Btl/Ac]"

'DrpColzDrsCC @BchRatHK
DrpCTT "$OHBch $BchRat $LasRat #3P"
End Sub

Private Sub W2TmpBchRat__Tst()
Dim A As Ymd: A = Ymd(20, 1, 30)
ClsAllCTbl
W2TmpOHBch A
W2TmpBchRat A
DoCmd.OpenTable "$BchRat"
End Sub
Private Sub W2TmpOHBch(A As Ymd): RunCQ W2TmpOHBchSql(A): End Sub
Private Function W2TmpOHBchSql$(A As Ymd)
W2TmpOHBchSql = FmtQQ("Select Distinct Sku,BchNo" & _
" into [$OHBch]" & _
" from [OH]" & _
" where ? and Btl>0 and Co=86", OHYmdBexp(A))
End Function
Private Sub W2TmpLasRat__Tst()
Dim A As Ymd: A = Ymd(20, 1, 30)
ClsAllCTbl
W2TmpOHBch A
W2TmpBchRat A
W2TmpLasRat A
DoCmd.OpenTable "$LasRat"
End Sub
Private Sub W2TmpBchRat(A As Ymd)
'Oup : $BchRat = Sku BchNo PermitD Rate Permit PermitDate

'Crt #A : PermitD Sku BchNo
RunCQ "Select Distinct x.Sku,x.BchNo,Max(x.PermitD) as PermitD" & _
" Into [#A]" & _
" From (PermitD x" & _
" Inner Join [$OHBch] a on a.BchNo=x.BchNo and a.Sku=x.Sku)" & _
" group by x.Sku,x.BchNo"

'Crt @BchRat: PermitD Sku BchNo Permit PermitDate
RunCQ "Select x.Sku,x.BchNo,a.Rate,a.PermitD,a.Permit,PermitDate" & _
" Into [$BchRat]" & _
" From ([#A] x" & _
" Inner Join [PermitD] a on a.PermitD=x.PermitD)" & _
" Inner Join [Permit]  b on a.Permit=b.Permit"
'Drp
DrpCT "#A"
End Sub
Private Sub W2TmpLasRat(A As Ymd)
'Oup: Sku BchNo LasPermit LasPermitDate LasBchNo
'Inp: $OHBch  : Sku BchNo
'Inp: $BchRat : Sku BchNo ...

'Crt #OHLasBch : Sku ! These are Sku+Bch using *Las
RunCQ "Select Distinct x.Sku" & _
" into [#OHLasBch]" & _
" from [$OHBch] x" & _
" left join [$BchRat] a on a.Sku=x.Sku and a.BchNo=x.BchNo" & _
" where a.Sku is null"

'Crt #A : PermitD Sku
RunCQ "Select Distinct x.Sku,Max(x.PermitD) as PermitD" & _
" Into [#A]" & _
" From (PermitD x" & _
" Inner Join [$OHBch] a on a.Sku=x.Sku)" & _
" group by x.Sku"

'Crt @BchRat: PermitD Sku BchNo Permit PermitDate
RunCQ "Select x.Sku,Rate,x.PermitD,a.Permit,PermitDate,a.BchNo" & _
" Into [$LasRat]" & _
" From ([#A] x" & _
" Inner Join [PermitD] a on a.PermitD=x.PermitD)" & _
" Inner Join [Permit]  b on a.Permit=b.Permit"

DrpCTT "#OHLasBch #A"
End Sub

Private Sub VVOupPacRat__Tst()
ClsAllCTbl
VVOupPacRat Ymd(20, 1, 30)
DoCmd.OpenTable "@PacRatD"
DoCmd.OpenTable "@PacRat"
End Sub
Private Function VVOupPacRat(A As Ymd) ' Crt @PacRatD @PacRat
'Inp: Tb-SkuRepackMulti = SkuNew SkuFm FmSkuQty
'Inp: Tb-OH             = Sku Co Btl ..
'Oup: @PacRatD   = Sku Des FmSku FmDes FmQty FmSkuAcRat RefPermit RefBchNo
'Oup: @PacRat    = Sku Des PacRat FmSkuCnt (PacRat is /Ac)

Dim Wh$: Wh = WhOHYmd(A)
'Crt #WiOHParSku : Sku
'    ===========
RunCQ FmtQQ("Select Distinct Sku as NewSku" & _
" into [#WiOHParSku]" & _
" from [OH]" & _
" ? and Btl>0 and Co=86 and Sku in (Select Distinct SkuNew from SkuRepackMulti)", Wh)

'Crt #ParChdSku : Sku FmSku | FmQty
'    ==========
RunCQ "Select SkuNew as NewSku, SkuFm as FmSku, FmSkuQty as FmQty" & _
" into [#ParChdSku]" & _
" from [SkuRepackMulti] x" & _
" where x.SkuNew in (Select NewSku from [#WiOHParSku])"

'Crt #FmSkuPmi1 : FmSku | Date_Id
'    ==========
RunCQ "Select Sku as FmSku,Max(Format(PermitDate,'YYYY-MM-DD') & '_' & x.Permit) as Date_Id" & _
" Into [#FmSkuPmi1]" & _
" From PermitD x" & _
" Inner join Permit a on x.Permit=a.Permit" & _
" where Sku in (Select FmSku from [#ParChdSku])" & _
" Group By Sku"

'Crt #FmSkuPmi2 : FmSku | PermitDate Permit
'    ==========
RunCQ "Select FmSku,CDate(Left(Date_Id,10)) as PermitDate, CLng(Mid(Date_Id,12)) as Permit" & _
" Into [#FmSkuPmi2]" & _
" From [#FmSkuPmi1]"

'Crt #FmSkuRat : FmSku | PermitDate Permit Rate
'    =========
RunCQ "Select x.FmSku,x.Permit,x.PermitDate,Rate" & _
" into [#FmSkuRat]" & _
" from ([#FmSkuPmi2] x" & _
" inner join [PermitD] a on a.Sku=x.FmSku and a.Permit=x.Permit)"

'Crt @PacRatD : Sku FmSku Permit Rate
'    ========
RunCQ "Select x.NewSku,'' As Des,x.FmSku,'' As FmSkuDes,x.FmQty,Rate,Permit,PermitDate" & _
" Into [@PacRatD]" & _
" From [#ParChdSku] x" & _
" left join [#FmSkuRat] a on x.FmSku = a.FmSku "
RunCQ "Update [@PacRatD] x inner join Sku a on a.Sku=x.NewSku set x.Des=a.SkuDes"
RunCQ "Update [@PacRatD] x inner join Sku a on a.Sku=x.FmSku set x.FmSkuDes=a.SkuDes"

'Crt @PacRat : SkuNew Rate
'    ========
RunCQ "Select Distinct x.NewSku,'' as Des,Sum(FmQty*x.Rate) as Rate,Count(*) as FmSkuCnt" & _
" Into [@PacRat]" & _
" From [@PacRatD] x" & _
" Group by NewSku"
RunCQ "Update [@PacRat] x inner join Sku a on a.Sku=x.NewSku set x.Des=a.SkuDes"

'-- Stp: Drp
DrpCTT "#WiOHParSku #ParChdSku #FmSkuRat #FmSkuPmi1 #FmSkuPmi2"
End Function

'------------------------------------------------------------------------------------------------------------------------ VVOupBchBatMo
Private Sub VVOupBchRatMO__Tst()
VVOupBchRatMO WWSampYmd
DoCmd.OpenTable "@BchRatMO"
End Sub
Private Sub VVOupBchRatMO(A As Ymd)
'Oup: [@BchRatMO]: Sku BchNo Val Btl | IsZHT0 [HKD/MOP] [MOP/Litre] [HKD/XXX] [XXX] [XXX/Btl]   with supporting....
'Fm   [OH]         Sku BchNo Val Btl       where Wh and Co=87
'Fm   [Sku]        Sku TaxRateMO          ! TaxRateMO>0
'## Stp    Oup       Stru
' 1 TmpOH  #OH       From [OH]  for those Wh and Co=87
' 2 TmpSku #Sku      From [Sku] for those TaxRateMO>0
' 3 TmpMO  @BchRatMO
'Tmp  [@BchRatMO]
DoCmd.SetWarnings False
StsQry "BchRatMO"

Dim Wh$: Wh = WhOHYmd(A)

'-- 1 Stp-TmpOH ===========================================================
RunCQ "Select Sku,BchNo,Val,Btl into [#OH] from OH" & Wh & " and Co=87"
Stop
RunCQ "Alter Table [#OH] add column " & _
"IsZHT0      yesno," & _
"[HKD/MOP]   Double," & _
"[MOP/Litre] Currency," & _
"[HKD/XXX]   Double," & _
"[XXX]       Text(3)," & _
"[XXX/Btl]   Currency "

RunCQ "Update [#OH] x,MacauRatePrm a set " & _
"  IsZHT0     =False," & _
"x.[HKD/MOP]  =a.[HKD/MOP]," & _
"x.[MOP/Litre]=a.[MOP/Litre]," & _
"x.[HKD/XXX]  =a.[HKD/XXX]," & _
"x.[XXX]      =a.[XXX]"

'-- 2 Stp-TmpSku ========================================================
RunCQ "Select Sku into [#Sku] from Sku where Nz(TaxRateMO,0)<>0"
RunCQ "Update [#OH] x inner join [#Sku] a on a.Sku=x.Sku set x.IsZHT0=true"
RunCQ "Update [#OH] x inner join MacauOverRideRate a on a.Sku=x.Sku and a.BchNo=x.BchNo set x.[XXX/Btl]=a.[XXX/Btl]"

'-- 3 Stp-OupBchRatMO ==================================================
RunCQ "Select * into [@BchRatMO] from [#OH] where IsZHT0 or (not [XXX/Btl] is null)"
RunCQ "Alter Table [@BchRatMO] add column [Btl/AC] Integer,[Litre/Btl] Double,Litre Double,LitreHKD Currency,[10%A] Currency,[10%B] Currency,[10%HKD] Currency,HKD Currency,[HKD/AC] Double,BchRat Double"
RunCQ "Update [@BchRatMO] x inner join [Sku] a on a.Sku=x.Sku set x.[Litre/Btl]=a.[Litre/Btl],x.[Btl/AC]=a.[Btl/AC]"
'
RunCQ "Update [@BchRatMO] set Litre    = Btl * [Litre/Btl]"
RunCQ "Update [@BchRatMO] set LitreHKD = Litre * [MOP/Litre] / [HKD/MOP]"
RunCQ "Update [@BchRatMO] set [10%A]   = [XXX/Btl] * Btl * 0.1"
RunCQ "Update [@BchRatMO] set [10%B]   = [Val] * 0.1"
RunCQ "Update [@BchRatMO] set [10%HKD] = IIF(Isnull([XXX/Btl]),[10%B],[10%A])"
RunCQ "Update [@BchRatMO] Set HKD      = [LitreHKD] + [10%HKD]"
RunCQ "Update [@BchRatMO] Set [HKD/AC] = HKD/Btl*[Btl/AC]"
RunCQ "Update [@BchRatMO] SEt BchRat   = [HKD/AC]"
RunCQ "Alter Table [@BchRatMO] add column SkuDes Text(50)"
RunCQ "Update [@BchRatMO] x inner join Sku a on x.SKu=a.Sku set x.SkuDes = a.SkuDes"

'-- Rename [@BchRatMO] 4 fields with [Btl] into [Bott] due to template is using [Bott]
RenCFlds "@BchRatMO", _
    "Btl  XXX/Btl  Btl/AC  Litre/Btl", _
    "Bott XXX/Bott Bott/AC Size"

'== 4 Stp-DrpTmp
DrpCTT "#Sku #OH"
End Sub
'------------------------------------------------------------------------------------------------------------------------ VVOupZHT0Rat
Private Sub VVOupZHT0Rat()
'Oup: @ZHT0Rat (Co Sku  SkuDes   Uom       [Rate/Uom] [Btl/AC] [Unit/AC] ==> [Rate/AC] [ZHT0Rat]
'Fm:  Sku       x  x    SkuDes   TaxUomHK  TaxRateHK   x        x
'                         SkuDes   TaxUomMO  TaxRateMO                                           <== TaxUomHK is in HKD/Uom
'     Calc                                                                      xx     xx        <== They are same
'               ZHT0Rat = Rate/Uom

'-- @ZHT0Rat: 7 Fields from Sku
DoCmd.SetWarnings False
Sts "ZHT0Rat"
Dim A$: A = "Insert into [@ZHT0Rat]" & _
                   " (Co,     Sku,SkuDes,            Uom,             [Rate/Uom],[Btl/AC],[Unit/AC]) "
RunCQ "Select 86 As Co,Sku,SkuDes,TaxUomHK as Uom,TaxRateHK As [Rate/Uom],[Btl/AC],[Unit/AC] into [@ZHT0Rat] From [SKU] where Nz(TaxRateHK,0)<>0"
RunCQ A & _
             "Select 87 As Co,Sku,SkuDes,TaxUomMO as Uom,TaxRateMO As [Rate/Uom],[Btl/AC],[Unit/AC]                 From [SKU] where Nz(TaxRateMO,0)<>0"

'-- @ZHT0Rat: Add 2 fields [Rate/AC] [ZHT0Rat]
RunCQ "Alter Table [@ZHT0Rat] add column [Rate/AC] Double, [ZHT0Rat] Double"

'-- @ZHT0Rat: Update [Rate/AC]
Select Case 1
Case 1
    'Not work:
    RunCQ "Update [@ZHT0Rat] set [Rate/AC]=[Rate/Uom]*[Unit/AC] Where Uom='PCE'" 'PCE is Set
    RunCQ "Update [@ZHT0Rat] set [Rate/AC]=[Rate/Uom]           Where Uom='CA'"  'CA  is AC
    RunCQ "Update [@ZHT0Rat] set [Rate/AC]=[Rate/Uom]*[Btl/AC]  Where Uom='COL'" 'COL is Btl
Case 2
    With CurrentDb.OpenRecordset("Select [Rate/AC],[Rate/Uom],[Btl/AC],[Unit/AC],Uom From [@ZHT0Rat] where Uom in ('PCE','CA','COL')")
        While Not .EOF
            .Edit
            Select Case !UOM
            Case "PCE": .Fields("Rate/AC").Value = .Fields("Rate/Uom").Value * .Fields("Unit/AC").Value
            Case "CA":  .Fields("Rate/AC").Value = .Fields("Rate/Uom").Value
            Case "COL": .Fields("Rate/AC").Value = .Fields("Rate/Uom").Value * .Fields("Btl/AC").Value
            End Select
            .Update
            .MoveNext
        Wend
    End With
End Select
RunCQ "Update [@ZHT0Rat] set [ZHT0Rat]=[Rate/AC]"
End Sub

'-- Gen & Fmt
Private Sub VVGenOFx__Tst():  VVGenOFx MB52LasOFx:                   End Sub

Private Function WWSampYmd() As Ymd: WWSampYmd = Ymd(20, 1, 30): End Function

'== Move to other module
Function GenOupFx(Fx$, Tp$, Fb$) As Workbook '#Gen-Oup-Fx# Cpy To to Fx and ref the Fx from @Fb and return saved Ws
CpyFfn Tp, Fx
Dim X As Excel.Application: Set X = NwXls
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(Fx)
RfhWb Wb, Fb
RfhWbPc Wb
Wb.Save
MaxvWb Wb
Set GenOupFx = Wb
End Function
