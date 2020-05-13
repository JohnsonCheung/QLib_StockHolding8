Attribute VB_Name = "gzLoadGit"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzLoadGit."
Sub GitFrm_RecordSource__Tst()
Dim A$: A = GitFrm_RecordSource(SampDDec24)
CurrentDb.OpenRecordset A
End Sub
Function GitFrm_RecordSource$(A As Ymd)
GitFrm_RecordSource$ = "SELECT x.*, PHNamDes, PHBrdDes, GITLoadDte, DateSerial(x.YY,x.MM,x.DD) as GITDte" & _
" FROM (GIT x" & _
" INNER JOIN PH_Brd a ON a.PHBrd = x.PHBrd)" & _
" INNER JOIN Report b ON x.YY=b.YY and x.MM=b.MM and x.DD=b.DD" & _
OHYmdBexp(A, "x.")
End Function

Private Sub LoadGit__Tst()
LoadGit SampDDec24
End Sub

Function IsCnlNoGit(A As Ymd) As Boolean
If HasGit(A) Then Exit Function
IsCnlNoGit = MsgBox("Git is not loaded!!" & vbCrLf & "[Ok] = Contine generate report" & vbCrLf & "[Cancel] = Cancel", vbOKCancel + vbQuestion) = vbCancel
End Function
Function HasGit(A As Ymd)
HasGit = HasRecCQ("Select * from Git" & OHYmdBexp(A))
End Function

Sub LoadGit(A As Ymd)
If Not Cfm("Start Load GIT?") Then Exit Sub

'Aim: Load IFx to GIT to GITDet & GIT.  Return True if loaded.
'Inp: >GIT = [Purchasing Document] Plant Material Currency [Still to be invoiced (qty)] [Still to be invoiced (val#)] HKD  ' [Still to be invoiced (qty)] is in ActCase
'Oup: GIT  = YY MM DD PHBrd | SC HKD DteCrt DteUpDD
'Oup: GITDet= YY MM DD Co Sku PoNo | Btl Ac Sc Amt Cur HKD
'Oup: Report  Update Report->(GITSc GITHKD)
'Ref: SKU     = SKU | PHBrd [Blt/AC] [Litre/Btl] [Btl/SC]
'Tmp: #LoadGIT_GITDet = Sku CdPlnt Co Ac Sc Btl Val
'Tmp: #LoadGIT_GIT
'Tmp: #LoadGIT_Report

'Logic:
'        Vdt:IFx
'        Crt:#GITDet
'
'Rpl: GITDet all records with Given date will be deleted and inserted (ie, replaced)
'Rpl: GIT    all records with Given date will be deleted and inserted (ie, replaced)
'Upd: Report = YY MM DD | GITSc GITHKD ..

'1 Vdt@IFx ========
Dim IFx$: IFx = GitIFx(A)
Const Wsn$ = "Sheet1"
Const FldNmCsv$ = "Plant,Purchasing Document,Material,Still to be invoiced (qty),Currency,Still to be invoiced (val#),HKD"
Const XlsTyCsv$ = "T    ,T                  ,T       ,N                         ,T       ,N                          ,N "
ChkWsCol IFx, Wsn, FldNmCsv, XlsTyCsv

CLnkFxw IFx, Wsn, ">GIT"
Dim mWhYmd$: mWhYmd = OHYmdBexp(A)
DoCmd.SetWarnings False
DrpCT "#LoadGIT"
'-------------------------------------------------------
'2.1
'Crt: #LoadGIT_GITDet
'Fm : >GIT
RunCQ "SELECT Distinct" & _
             " `Purchasing Document` as PoNo  ," & _
                                " '' as PHBrd ," & _
                          " Material AS Sku   ," & _
                             " Plant as CdPlnt," & _
                          " CByte(0) as Co    ," & _
 " Sum(`Still to be invoiced (qty)`) as Ac    ," & _
                           " CDbl(0) as Sc    ," & _
                           " CLng(0) as btl   ," & _
                          " Currency as Cur   ," & _
" Sum(`Still to be invoiced (val#)`) as Amt   ," & _
                        " Sum(x.HKD) as HKD" & _
" INTO [#LoadGIT_GITDet]" & _
" FROM [>GIT] x Where Nz(`Still to be invoiced (qty)`,0)<>0" & _
" Group By  `Purchasing Document`,Material,Plant,Currency"
RunCQ "UpDate `#LoadGIT_GITDet` x inner join Sku       a on a.Sku=x.Sku Set   Btl=Ac*[Btl/AC]"           ' UpDDate Btl
RunCQ "UpDate `#LoadGIT_GITDet` x inner join Sku       a on a.Sku=x.Sku Set PHBrd=Left(ProdHierarchy,4)" ' UpDDate Btl
RunCQ "UpDate `#LoadGIT_GITDet` x inner join qSku_Main a on a.Sku=x.Sku Set    Sc=Btl/[Btl/SC]"          ' UpDDate Sc
RunCQ "UpDate `#LoadGIT_GITDet` set Co=Left(CdPlnt,2)"                  ' Update Co
'-------------------------------------------------------
'2.2
'Rpl: GITDet
'By : #LoadGIT_GITDet
RunCQ "DELETE FROM GITDet" & mWhYmd
With A
RunCQ FmtStr("INSERT INTO GITDet (YY,MM,DD,Co,Sku,PoNo,Ac,Sc,Btl,Amt,Cur,HKD) Select {0},{1},{2},Co,Sku,PoNo,Ac,Sc,Btl,Amt,Cur,HKD from `#LoadGIT_GITDet`", .Y, .M, .D)
End With
'-------------------------------------------------------
'3.1
'Crt: #LoadGIT_GIT    = Sku CdPlnt Co Ac Sc Btl Val
'Fm : #LoadGIT_GITDet = Co PHBrd Sc HKD
'Ref: Sku             = Sku PHBrd ..
'#GIT   =PHBrd Sc HKD
'#GITDet=Sku Plant Sc Ac Cur Amt HKD
RunCQ "SELECT distinct Co,PHBrd,Sum(x.Sc) as Sc,Sum(x.HKD) as HKD" & _
" into `#LoadGIT_GIT`" & _
" from `#LoadGIT_GITDet` x left join Sku a on a.Sku=x.Sku" & _
" GROUP BY Co,PHBrd"
'---------------------------------------------------------
'3.2 Rpl: GIT
'By : #LoadGIT_GIT  ' Drop after use
Dim W$: W = OHYmdBexp(A)
RunCQ FmtStr("DELETE FROM GIT" & W)
With A
RunCQ FmtStr("INSERT INTO GIT (YY,MM,DD,PHBrd,Sc,HKD) Select {0},{1},{2},PHBrd,Sc,HKD from `#LoadGIT_GIT`", .Y, .M, .D)
End With
'---------------------------------------------------------
'4. Upd: Report->(GITSc GITHKD)
'By : #LoadGIT_GITDet               ' Drop after use
RunCQ "SELECT Count(*) As GITNRec,Sum(Sc) AS GITSc,Sum(Ac) as GITAc,Sum(Btl) as GITBtl,Sum(HKD) as GITHKD INTO [#LoadGIT_Report] FROM [#LoadGIT_GITDet]"
RunCQ "UpDATE Report x, [#LoadGIT_Report] a SET x.GITNRec=a.GITNRec,x.GITSc=a.GITSc, x.GITAc=a.GITAc, x.GITBtl=a.GITBtl, x.GITHKD=a.GITHKD, GITLoadDte=Now" & W

DrpCTT "#LoadGIT_Report #LoadGIT_GIT #LoadGIT_GITDet >GIT"
Done
End Sub
