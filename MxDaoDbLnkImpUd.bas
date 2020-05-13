Attribute VB_Name = "MxDaoDbLnkImpUd"
Option Explicit
Option Compare Text
Enum eLnkisFilTy: eFb: eFx: End Enum
'-- Src
Type LnkisInp: Ix As Integer: Inpn As String: Ffn As String: End Type 'Deriving(Ay Ctor)
Type LnkisFb: Ix As Integer: Inpn As String: Tny() As String: End Type 'Deriving(Ay Ctor)
Type LnkisFx: Ix As Integer: Inpn As String: Inpnw As String: Stru As String: End Type 'Deriving(Ay Ctor)
Type LnkisWh: Ix As Integer: Tbn As String: Bexp As String: End Type 'Deriving(Ay Ctor)
Type LnkisFld: Ix As Integer: Intn As String: Ty As String: Extn As String: End Type 'Deriving(Ay Ctor)
Type LnkisStru: Ix As Integer: Stru As String: Fld() As LnkisFld: End Type 'Deriving(Ay Ctor)
Type Lnkis
    Inp() As LnkisInp
    FbTbl() As LnkisFb
    FxTbl() As LnkisFx
    TblWh() As LnkisWh
    Stru() As LnkisStru
    MustHasRecTbl() As ILn
End Type

Function SampInpFilSrc() As String()
X "Inp"
X " DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X " ZHT0    C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X " MB52    C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X " Uom     C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X " GLBal   C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
SampInpFilSrc = XX
End Function

Function SampLnkImpSrc() As String()
Erase XX
X "*Spec :LnkImp MB52 *Inp- FbTbl- FxTbl- Tbl.Where- *Stru MusHasRecTbl"
X " Inp::(Inpn,Ffn)+"
X " FbTbl::(FbTbn,Stru..)*"
X " FxTbl::(FxTbn,?Inpnw,?Stru)*  "
X "         Inpnw is Inpnw is Inpn-dot-Wsn.  It is optional.  Inpn will use FxInpn and Wsn will use sheet1"
X " Tbl.Where::(Inpn,Bexp)*                  The Bexp is using Extn in Sql-Bexp"
X " Stru::(Stru,(Intn,?Ty,?Extn))+           "
X "          Ty is (Dbl | Txt Dbl|Txt Dte)"
X "          Extn is a term, must quoated in []"
X " MustHasRec::(Inpn..|*AllInp)"
X "          *AllInpn all Inpn should have record"
X "Inp"
X " DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X " ZHT0    C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X " MB52    C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X " Uom     C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X " GLBal   C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
X "FbTbl"
X " --  Fbn T.."
X " DutyPay Permit PermitD"
X "FxTbl "
X " -- FxTbn Inpnw    Stru"
X " ZHT086  ZHT0.8600 ZHT0"
X " ZHT087  ZHT0.8700 ZHT0"
X " MB52"
X " Uom"
X " GLBal"
X "Tbl.Where"
X " MB52 Plant='8601' and [Storage Location] in ('0002','')"
X " Uom  Plant='8601'"
X "Stru Permit"
X " Permit"
X " PermitNo"
X " PermitDate"
X " PostDate"
X " NSku"
X " Qty"
X " Tot"
X " GLAc"
X " GLAcName"
X " BankCode"
X " ByUsr"
X " DteCrt"
X " DteUpd"
X "Stru PermitD"
X " PermitD"
X " Permit"
X " Sku"
X " SeqNo"
X " Qty"
X " BchNo"
X " Rate"
X " Amt"
X " DteCrt"
X " DteUpd"
X "Stru.ZHT0"
X " Sku       Txt Material    "
X " CurRateAc Dbl [     Amount]"
X " VdtFm     Txt Valid From  "
X " VdtTo     Txt Valid to    "
X " HKD       Txt Unit        "
X " Per       Txt per         "
X " CA_Uom    Txt Uom         "
X "Stru MB52"
X " Sku1    Txt Material          "
X " Sku    Txt Material          "
X " Whs    Txt Plant             "
X " Loc    Txt Storage Location  "
X " BchNo  Txt Batch             "
X " QInsp  Dbl In Quality Insp#  "
X " QUnRes Dbl UnRestricted      "
X " QBlk   Dbl Blocked           "
X " VInsp  Dbl Value in QualInsp#"
X " VUnRes Dbl Value Unrestricted"
X " VBlk   Dbl Value BlockedStock"
X " VBlk2  Dbl Value BlockedStock1"
X " VBlk1  Dbl Value BlockedStock2"
X "Stru Uom"
X " Sc_U    Txt SC "
X " Topaz   Txt Topaz Code "
X " ProdH   Txt Product hierarchy"
X " Sku     Txt Material            "
X " Des     Txt Material Description"
X " AC_U    Txt Unit per case       "
X " SkuUom  Txt Base Unit of Measure"
X " BusArea Txt Business Area       "
X "Stru GLBal"
X " BusArea Txt Business Area Code"
X " GLBal   Dbl                   "
X "Stru SkuTaxBy3rdParty"
X " SkuTaxBy3rdParty "
X "Stru SkuNoLongerTax"
X " SkuNoLongerTax"
X "MustHasRecTbl"
X " *AllInp"
SampLnkImpSrc = XX
End Function

Function SampLnkImpSrc1() As String()
Erase XX
X "Inp"
X " DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X " ZHT0  C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X " MB52  C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X " Uom   C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X " GLBal C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
X "FbTbl"
X " --  Fbn Tn.."
X " DutyPay Permit PermitD"
X "FxTbl T  FxNm.Wsn  Stru"
X " ZHT086  ZHT0.8600 ZHT0"
X " ZHT087  ZHT0.8700 ZHT0"
X " MB52                  "
X " Uom                   "
X " GLBal"
X "Tbl.Where"
X " MB52 Plant='8601' and [Storage Location] in ('0002','')"
X " Uom  Plant='8601'"
X "Stru.Permit"
X " Permit"
X " PermitNo"
X " PermitDate"
X " PostDate"
X " NSku"
X " Qty"
X " Tot"
X " GLAc"
X " GLAcName"
X " BankCode"
X " ByUsr"
X " DteCrt"
X " DteUpd"
X "Stru.PermitD"
X " PermitD"
X " Permit"
X " Sku"
X " SeqNo"
X " Qty"
X " BchNo"
X " Rate"
X " Amt"
X " DteCrt"
X " DteUpd"
X "Stru.ZHT0"
X " Sku       Txt Material    "
X " CurRateAc Dbl [     Amount]"
X " VdtFm     Txt Valid From  "
X " VdtTo     Txt Valid to    "
X " HKD       Txt Unit        "
X " Per       Txt per         "
X " CA_Uom    Txt Uom         "
X "Stru.MB52"
X " Sku    Txt Material          "
X " Whs    Txt Plant             "
X " Loc    Txt Storage Location  "
X " BchNo  Txt Batch             "
X " QInsp  Dbl In Quality Insp#  "
X " QUnRes Dbl UnRestricted      "
X " QBlk   Dbl Blocked           "
X " VInsp  Dbl Value in QualInsp#"
X " VUnRes Dbl Value Unrestricted"
X " VBlk   Dbl Value BlockedStock"
X "Stru.Uom"
X " Sc_U    Txt SC "
X " Topaz   Txt Topaz Code "
X " ProdH   Txt Product hierarchy"
X " Sku     Txt Material            "
X " Des     Txt Material Description"
X " AC_U    Txt Unit per case       "
X " SkuUom  Txt Base Unit of Measure"
X " BusArea Txt Business Area       "
X "Stru.GLBal"
X " BusArea Txt Business Area Code"
X " GLBal   Dbl                   "
X "Stru.SkuRepackMulti"
X " SkuRepackMulti   GLBal   Dbl                     "
X "Stru.SkuTaxBy3rdParty"
X " SkuTaxBy3rdParty GLBal   Dbl                     "
X "Stru.SkuNoLongerTax"
X " SkuNoLongerTax"
SampLnkImpSrc1 = XX
Erase XX
End Function




Function LnkisInp(Ix, Inpn, Ffn) As LnkisInp
With LnkisInp
    .Ix = Ix
    .Inpn = Inpn
    .Ffn = Ffn
End With
End Function
Function AddLnkisInp(A As LnkisInp, B As LnkisInp) As LnkisInp(): PushLnkisInp AddLnkisInp, A: PushLnkisInp AddLnkisInp, B: End Function
Sub PushLnkisInpAy(O() As LnkisInp, A() As LnkisInp): Dim J&: For J = 0 To LnkisInpUB(A): PushLnkisInp O, A(J): Next: End Sub
Sub PushLnkisInp(O() As LnkisInp, M As LnkisInp): Dim N&: N = LnkisInpSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkisInpSi&(A() As LnkisInp): On Error Resume Next: LnkisInpSi = UBound(A) + 1: End Function
Function LnkisInpUB&(A() As LnkisInp): LnkisInpUB = LnkisInpSi(A) - 1: End Function
Function LnkisFb(Ix, Inpn, Tny$()) As LnkisFb
With LnkisFb
    .Ix = Ix
    .Inpn = Inpn
    .Tny = Tny
End With
End Function
Function AddLnkisFb(A As LnkisFb, B As LnkisFb) As LnkisFb(): PushLnkisFb AddLnkisFb, A: PushLnkisFb AddLnkisFb, B: End Function
Sub PushLnkisFbAy(O() As LnkisFb, A() As LnkisFb): Dim J&: For J = 0 To LnkisFbUB(A): PushLnkisFb O, A(J): Next: End Sub
Sub PushLnkisFb(O() As LnkisFb, M As LnkisFb): Dim N&: N = LnkisFbSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkisFbSi&(A() As LnkisFb): On Error Resume Next: LnkisFbSi = UBound(A) + 1: End Function
Function LnkisFbUB&(A() As LnkisFb): LnkisFbUB = LnkisFbSi(A) - 1: End Function
Function LnkisFx(Ix, Inpn, Inpnw, Stru) As LnkisFx
With LnkisFx
    .Ix = Ix
    .Inpn = Inpn
    .Inpnw = Inpnw
    .Stru = Stru
End With
End Function
Function AddLnkisFx(A As LnkisFx, B As LnkisFx) As LnkisFx(): PushLnkisFx AddLnkisFx, A: PushLnkisFx AddLnkisFx, B: End Function
Sub PushLnkisFxAy(O() As LnkisFx, A() As LnkisFx): Dim J&: For J = 0 To LnkisFxUB(A): PushLnkisFx O, A(J): Next: End Sub
Sub PushLnkisFx(O() As LnkisFx, M As LnkisFx): Dim N&: N = LnkisFxSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkisFxSi&(A() As LnkisFx): On Error Resume Next: LnkisFxSi = UBound(A) + 1: End Function
Function LnkisFxUB&(A() As LnkisFx): LnkisFxUB = LnkisFxSi(A) - 1: End Function
Function LnkisWh(Ix, Tbn, Bexp) As LnkisWh
With LnkisWh
    .Ix = Ix
    .Tbn = Tbn
    .Bexp = Bexp
End With
End Function
Function AddLnkisWh(A As LnkisWh, B As LnkisWh) As LnkisWh(): PushLnkisWh AddLnkisWh, A: PushLnkisWh AddLnkisWh, B: End Function
Sub PushLnkisWhAy(O() As LnkisWh, A() As LnkisWh): Dim J&: For J = 0 To LnkisWhUB(A): PushLnkisWh O, A(J): Next: End Sub
Sub PushLnkisWh(O() As LnkisWh, M As LnkisWh): Dim N&: N = LnkisWhSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkisWhSi&(A() As LnkisWh): On Error Resume Next: LnkisWhSi = UBound(A) + 1: End Function
Function LnkisWhUB&(A() As LnkisWh): LnkisWhUB = LnkisWhSi(A) - 1: End Function
Function LnkisFld(Ix, Intn, Ty$, Extn) As LnkisFld
With LnkisFld
    .Ix = Ix
    .Intn = Intn
    .Ty = Ty
    .Extn = Extn
End With
End Function
Function AddLnkisFld(A As LnkisFld, B As LnkisFld) As LnkisFld(): PushLnkisFld AddLnkisFld, A: PushLnkisFld AddLnkisFld, B: End Function
Sub PushLnkisFldAy(O() As LnkisFld, A() As LnkisFld): Dim J&: For J = 0 To LnkisFldUB(A): PushLnkisFld O, A(J): Next: End Sub
Sub PushLnkisFld(O() As LnkisFld, M As LnkisFld): Dim N&: N = LnkisFldSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkisFldSi&(A() As LnkisFld): On Error Resume Next: LnkisFldSi = UBound(A) + 1: End Function
Function LnkisFldUB&(A() As LnkisFld): LnkisFldUB = LnkisFldSi(A) - 1: End Function
Function LnkisStru(Ix, Stru, Fld() As LnkisFld) As LnkisStru
With LnkisStru
    .Ix = Ix
    .Stru = Stru
    .Fld = Fld
End With
End Function
Function AddLnkisStru(A As LnkisStru, B As LnkisStru) As LnkisStru(): PushLnkisStru AddLnkisStru, A: PushLnkisStru AddLnkisStru, B: End Function
Sub PushLnkisStruAy(O() As LnkisStru, A() As LnkisStru): Dim J&: For J = 0 To LnkisStruUB(A): PushLnkisStru O, A(J): Next: End Sub
Sub PushLnkisStru(O() As LnkisStru, M As LnkisStru): Dim N&: N = LnkisStruSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkisStruSi&(A() As LnkisStru): On Error Resume Next: LnkisStruSi = UBound(A) + 1: End Function
Function LnkisStruUB&(A() As LnkisStru): LnkisStruUB = LnkisStruSi(A) - 1: End Function
