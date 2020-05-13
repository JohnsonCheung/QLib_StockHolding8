Attribute VB_Name = "MxCliMHDRelCst"
Option Compare Text
Option Explicit
Const CNs$ = "App.ShpCst"
Const CLib$ = "QRelCst."
Const CMod$ = CLib & "MxCliMHDRelCst."
Function RelCstPgmDb() As Database: Set RelCstPgmDb = Db(RelCstFba): End Function

Sub UomDoc()
#If False Then
InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
Oup : UOM        Sku      SkuUOM                 Des                    Sc_U

Note on [Sales text.xls]
Col  Xls Title            FldName     Means
F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
J    Unit per case        Sc_U        how many unit per AC
K    SC                   SC_U        how many unit per SC   ('no need)
L    COL per case         AC_B        how many bottle per AC
-----
Letter meaning
B = Bottle
AC = act case
SC = standard case
U = Unit  (Bottle(COL) or Set (PCE))

 "SC              as SC_U," & _  no need
 "[COL per case]  as AC_B," & _ no need
#End If
End Sub

Function Plant8687MisEr(FxMB52$, Wsn$) As String()
If NReczFxw(FxMB52, Wsn, "Plant in ('8601','8701')") = 0 Then
    Plant8687MisEr = W1Msg(FxMB52, Wsn)
End If
End Function
Private Function W1Msg(FxMB52$, Wsn$) As String()
Const M$ = "Column-[Plant] must have value 8601 or 8701"
W1Msg = FmtFmsgNap("Plant8687MisEr", M, "MB52-File Worksheet", FxMB52, Wsn)
End Function

Private Sub ROup(D As Database)
ORate D
OMain D
End Sub
Private Function OMain$(D As Database)
'#IUom
'#IMB52
'@IMB52 :Drs-Whs-Sku-QUnRes-QBlk-QInsp
'@IUom  :Sku-Sc_U-Des-StkUom
'Ret      : @@
DrpT D, "@Main"

'== Crt @Main fm #IMB52
'   Whs Sku OH Des StkUom Sc_U OH
RunQ D, "Select Distinct Whs,Sku,Sum(QUnRes+QBlk+QInsp) As OH into [@Main] from [#IMB52] Group by Whs,Sku"
RunQ D, "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10),Sc_U Int, OH_Sc Double"
RunQ D, "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"
RunQ D, "Update [@Main] set OH_Sc=OH/Sc_U where Sc_U>0"

'== Add Col Stream ProdH F2 M32 M35 M37 Topaz ZHT1 RateSc Z2 Z5 Z7
'   Upd Col ProdH Topaz
'   Upd Col F2 M32 M35 M37
RunQ D, "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH text(7), F2 Text(2), M32 text(2), M35 text(5), M37 text(7), ZHT1 Text(7), Z2 text(2), Z5 text(5), Z7 text(7), RateSc Currency, Amt Currency"
RunQ D, "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"
RunQ D, "Update [@Main] set F2=Left(ProdH,2),M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M37=Mid(ProdH,3,7)"

'== Upd Col ZHT1 RateSc
RunQ D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M37=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
RunQ D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
RunQ D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'Stream
'Z2 Z5 Z7
'Amt
RunQ D, "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"
RunQ D, "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z7=Left(ZHT1,7) where not ZHT1 is null"
RunQ D, "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
End Function
Private Function ORate$(D As Database)
'VdtFm & VdtTo format DD.MM.YYYY
'1: #IZHT18701 VdtFm VdtTo L3 RateSc
'1: #IZHT18601 VdtFm VdtTo L3 RateSc
'2: #IUom     SKu Sc_U
'O: @Rate  ZHT1 RateSc
DrpTT D, "#Cpy1 #Cpy2 #Cpy @Rate"
RunQ D, "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
RunQ D, "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

RunQ D, "Select * into [#Cpy] from [#Cpy1] where False"
RunQ D, "Insert into [#Cpy] select * from [#Cpy1]"
RunQ D, "Insert into [#Cpy] select * from [#Cpy2]"

RunQ D, "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
RunQ D, "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
RunQ D, "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

RunQ D, "Select Whs,ZHT1,RateSc into [@Rate] from [#Cpy]"
DrpTT D, "#Cpy #Cpy1 #Cpy2"
End Function

