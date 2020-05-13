Attribute VB_Name = "gzLoadMB52"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzLoadMB52."
Option Base 0
Public Const MB52Wsn$ = "Sheet1"
'Public Const MB52IFxFF$ = "  Material Plant [Storage Location] Batch [Base Unit of Measure] Unrestricted [Transit and Transfer] [In Quality Insp#] Blocked [Value Unrestricted] [Val# in Trans#/Tfr] [Value in QualInsp#] [Value BlockedStock] [Value Rets Blocked]"
'Public Const MB52IFxTyFF$ = "T        T      T                 T      T                     N             N                      N                 N        N                    N                    N                    N                    N"

Public Const MB52FldNmCsv$ = "Material,Plant,Storage Location,Batch,Base Unit of Measure,Unrestricted,Transit and Transfer,In Quality Insp#,Blocked,Value Unrestricted,Val# in Trans#/Tfr,Value in QualInsp#,Value BlockedStock,Value Rets Blocked"
Public Const MB52FldTyCsv$ = "T       ,T    ,T               ,T    ,T                   ,N           ,N                   ,N               ,N      ,N                 ,N                 ,N                 ,N                 ,N      "


Sub T_LoadMB52(): LoadMB52__Tst: End Sub
'---
Private Sub LoadMB52__Tst()
LoadMB52 Ymd(20, 1, 30), NoAsk:=True
End Sub

Sub LoadMB52(A As Ymd, Optional NoAsk As Boolean)
Dim IFx$: IFx = MB52IFx(A)
ChkFfnExist IFx
If Not Cfm("Start Load MB52?", , NoAsk) Then Exit Sub
Sts "Start loading MB52: " & YYmdStr(A)
YChk A
LnkMB52 A
YImp
Dim Wh$: Wh = OHYmdBexp(A)
DoCmd.SetWarnings False
'---------------------------------------------------------------------------
'Crt : #OH
'Fm  : #IMB52         ' Drop after use
RunCQ "SELECT Co, SKU, BchNo, Sum(x.Q) AS Q, Sum(x.V) AS V, CByte(0) as YpStk, SLoc" & _
" INTO [#OH]" & _
" FROM [#IMB52] x" & _
" GROUP BY Co,SKU,BchNo,SLoc;"
RunCQ "DELETE * FROM [#OH] WHERE Nz(Q,0)=0 AND Nz(V,0)=0;"
RunCQ "Update [#OH] x inner join YpStk a on a.Co=x.Co and a.SLoc=x.SLoc set x.YpStk=a.YpStk"
RunCQ "Drop Table `#IMB52`"
'---------------------------------------------------------------------------
'Update Q -> Btl  (Q is in StkUnit, which may PCE or COL  (PCE is set.  Required to convert to Btl)
RunCQ "Alter Table [#OH] add column Btl Long,[Unit/AC] double, [Btl/AC] integer, [Unit/SC] double"
RunCQ "Update [#OH] x inner join [qSku_Main] a on x.Sku=a.Sku set " & _
"x.[Unit/SC]=a.[Unit/SC]," & _
"x.[Btl/AC]=a.[Btl/AC]," & _
"x.[Unit/Ac]=a.[Unit/Ac]"
RunCQ "Update [#OH] set Btl = Q / [Unit/AC] * [Btl/AC]"
'---------------------------------------------------------------------------
'Upd: OH
'Fm : #OH
RunCQ "DELETE FROM OH" & Wh
With A
RunCQ FmtStr("INSERT INTO OH (YY, MM, DD, Co, SLoc, SKU, YpStk, BchNo, Btl, Val )" & _
" SELECT {0}, {1}, {2}, Co, SLoc, SKU, YpStk, BchNo, Btl, V FROM [#OH];", .Y, .M, .D)
End With
'---------------------------------------------------------------------------
'Upd: Report->(NRecMB52 TotBtlMB52 TotHKDMB52)
'Fm : #OH                                  ' Drop after use
RunCQ "SELECT Count(*) AS NRec, Sum(Q) AS TotBtl, Sum(V) AS TotVal INTO [#Tot] FROM [#OH]"
RunCQ "UPDATE Report x, [#Tot] a SET DteMB52=Now(), x.NRecMB52=a.NRec, x.TotBtlMB52=a.TotBtl, x.TotHKDMB52=a.TotVal" & Wh
RunCQ "Drop Table `#Tot`"

'Upd: Report->(MB52SC MB52AC)
RunCQ "Select Distinct Sku,Sum(Q) as Btl into [#A] from [#OH] Group by Sku"
RunCQ "Alter Table [#A] add column [Unit/SC] double, [Unit/AC] double, [Btl/AC] double, SC double,AC double"
RunCQ "Update [#A] x inner join qSku_Main a on a.Sku=x.Sku set x.[Unit/SC]=a.[Unit/SC],x.[Unit/AC] = a.[Unit/AC], x.[Btl/AC]=a.[Btl/AC]"
RunCQ "Update [#A] set AC=Btl/[Btl/AC], SC==Btl/[Btl/AC] * [Unit/AC] / [Unit/SC]"
RunCQ "Select Sum(x.SC) as SC, Sum(x.AC) as AC into [#B] from [#A] x"
'DoCmd.RunSQL "Update Report,[#B] set MB52SC=SC,MB52AC=AC," & _
'"GitSC=Null,GitAC=Null,GitHKD=Null,GitBtl=Null,GitNRec=Null,GitLoadDte=Null" & Wh
RunCQ "Update Report,[#B] set MB52SC=SC,MB52AC=AC" & Wh

'-- Drp temp table
RunCQ "Drop Table [#A]"
RunCQ "Drop Table [#B]"
RunCQ "Drop Table [#OH]"
Done
End Sub

'== Y
Private Sub YImp()
Sts "Importing....."
RunCQ "SELECT Material AS SKU, Batch as BchNo, CByte(Left(Plant,2)) as Co, [Storage Location] as SLoc," & _
" CDbl(Nz(Unrestricted,0)+ Nz([Transit and Transfer],0)+Nz([In Quality Insp#],0)+Nz(Blocked,0)) As Q," & _
" CCur(Nz([Value Unrestricted],0)+Nz([Val# in Trans#/Tfr],0)+Nz([Value in QualInsp#],0)+Nz([Value BlockedStock],0)) As V" & _
" INTO [#IMB52]" & _
" FROM [>MB52];"
End Sub
'--
Private Sub YChk(A As Ymd)
Sts "Validating...":
Dim Fx$, W$
    Fx = MB52IFx(A)
    W = MB52Wsn
ChkWsCol Fx, W, MB52FldNmCsv, MB52FldTyCsv
W1ChkWsRec Fx, W
End Sub

Private Sub W1ChkWsRec(Fx$, W$)
Dim Plnt$()
    Plnt = DisSyzFxwc(Fx, W, "Plant")
    
Dim M1$, M2$()
    M1 = ShdNBEmsglzForFxwc(Fx, W, "Material")
    M2 = ShdAllInEmsgzForSy(Plnt, "8701", "8601")
Dim M$(): M = AddItmAy(M1, M2)
ChkEr M, BoxlzFxw(Fx, W, "There is errors in Excel file")
End Sub

'== To be move to Other module
Function BoxlzFxw$(Fx$, W$, BoxCxt$)
Dim B$
If BoxCxt <> "" Then B = Boxl(BoxCxt) & vbCrLf
BoxlzFxw = B & W2Fxwl(Fx, W)
End Function

Private Function W2Fxwl$(Fx$, W$)
Dim O$()
PushI O, "Excel file : " & QuoSq(Fn(Fx))
PushI O, "Path       : " & QuoSq(Pth(Fx))
PushI O, "Worksheet  : " & QuoSq(W)
W2Fxwl = JnCrLf(O)
End Function
