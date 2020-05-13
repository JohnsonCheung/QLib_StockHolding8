Attribute VB_Name = "gzLoadFc"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzLoadFc."
Option Base 0
Const MhWsn$ = "Table"

Const FcTy = "Final Forecast M-1"
'-- If without Co, it is Mh
Const T_LFc$ = ">Fc"
Const T_TIFc$ = "#IFc"
'-- If with Co, it is Ud
Const T_LFc86$ = ">Fc86"
Const T_TIFc86$ = "#IFc86"
Const T_LFc87$ = ">Fc87"
Const T_TIFc87$ = "#IFc87"
Const T_TFc$ = "#Fc"
Const Wsn86$ = "MhdHK - Std Case"
Const Wsn87$ = "MhdMO - Std Case"

Sub LoadFc(A As StmYM)
ChkFfnExist IFx(A), "Forecast"
Select Case A.Stm
Case "M"
    RMhWFx A
    RMhLnk A
    RMhImp
    RMhTmpFcSku
    RMhChk
Case "U"
    RUdWFx A
    RUdLnk A
    RUdImp
    RUdTmpFcSku
    RUdChk
End Select

UpdTbPHStkDays7 A    ' Load #IFcSku into FcSku   FcSku   = VerYY VerMM Co Sku M01..M15
RfhTbFc_Fm_TblFcSku A
ClrCAcsSts
DrpCTT "#Fc #IFc >Fc #IFc86 #IFc87 >Fc86 >Fc87"
If IsFrmOpn("LoadFc") Then Form_LoadFc.Requery
Done
End Sub
'---==============================================================================================
Private Sub RMhWFx__Tst()
RMhWFx LasMHFc
End Sub
Private Sub RMhWFx(A As StmYM)
'Inp: FcIFx(@A) for MD.  All Sku be belong UD
'Oup: Fc_WrkFx(@A) with sheet FcSku is added
'Ret : True if any err.  Prompt the err

'== Kill & Copy
Dim W$: W = FcWFx(A)
Dim I$: I = IFx(A)
DltFfnIf W
FileCopy I, W

'== Open
Dim Wb As Workbook: Set Wb = WbzFx(W)

'== Delet Sheet
Dim J%, Ws As Worksheet
For J = Wb.Sheets.Count To 1 Step -1
    Set Ws = Wb.Sheets(J)
    If Ws.Name <> "Table" Then Ws.Visible = xlSheetVisible: Ws.Delete
Next
'== Delete Row and Col
Set Ws = Wb.Sheets("Table")
Ws.Range("$1:$14").EntireRow.Delete
Ws.Range("$A:$F").EntireColumn.Delete
Ws.Range("$C:$E").EntireColumn.Delete
Ws.Range("$D:$D").EntireColumn.Delete
Ws.Range("$E:$E").EntireColumn.Delete

'== Delete Shapes
While Ws.Shapes.Count > 0
    Ws.Shapes(1).Delete
Wend
'== Column D == FcTy
Ws.Range("D1").Value = "FcTy"
'== Insert Listobject
Dim C2R&: C2R = Ws.Cells.SpecialCells(xlCellTypeLastCell).Row
Dim C1 As Range: Set C1 = Ws.Range("A1")
Dim C2 As Range: Set C2 = Ws.Cells(C2R, "P")
Dim Rg As Range: Set Rg = Ws.Range(C1, C2)
Dim Lo As ListObject: Set Lo = Ws.ListObjects.Add(xlSrcRange, Rg, , xlYes)
Lo.TableStyle = "TableStyleLight1"
Lo.Range.AutoFilter Field:=4, Criteria1:=FcTy

'== ErCellValMsg
Dim O$()
Dim MNy$(): MNy = MonthNy(A.Y, A.M)
PushNB O, ErCellValMsg(Ws, "A1", "Market")
PushNB O, ErCellValMsg(Ws, "B1", "Market Channel")
PushNB O, ErCellValMsg(Ws, "C1", "Product")
PushNB O, ErCellValMsg(Ws, "D1", "FcTy")

For J = 0 To 11
    Dim Adr$: Adr = WsRC(Ws, 1, 5 + J).Address(False, False)
    PushNB O, ErCellValMsg(Ws, Adr, MNy(J))
Next

If Si(O) Then
    MaxiWb Wb
    Wb.Save
    Const L1$ = "Errors found in the working file."
    Const L2$ = "Fix the original Forecast file and load again"
    Const L3$ = "============================================="
    BrwEr AddSy(SyzAp(L1, L2, L3, ""), O)
End If

'== RenSku
'== RenCoNm
Set Rg = Ws.Range("C1"): Rg.Value = "Sku"
Set Rg = Ws.Range("B1"): Rg.Value = "CoNm"
'== Ren M01..12
For J = 1 To 12
    Set Rg = Ws.Cells(1, 4 + J)
    Rg.Value = "M" & Format(J, "00")
Next

'== Save / Close / Quit Xls
SavWbQuit Wb
End Sub
'---=================================================================================================================
Private Sub RUdWFx__Tst()
RUdWFx LasUDFc
End Sub
Private Sub RUdWFx(A As StmYM)
'Inp: FcFx for UD.  All Sku be belong UD
'Oup: $FcSku (Same stru as FcSku: VerYY,VerMM,YY,MM,Co,Sku,SC.  Verify all SKU should be UD, anything wrong report and raise error

'== Kill and Copy to WFx
Dim W$: W = FcWFx(A)
Dim I$: I = IFx(A)
DltFfnIf W
FileCopy I, W
'== Open
Dim X As Excel.Application: Set X = NwXls
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(W)
RUdWs Wb, 86, A.Y, A.M
RUdWs Wb, 87, A.Y, A.M
Wb.Close True
X.Quit
End Sub
Private Sub RUdWs(Wb As Workbook, Co As Byte, Y As Byte, M As Byte)
Dim Ws As Worksheet: Set Ws = Wb.Sheets(UdWsn(Co))
'== Delete Row and Col
Ws.Range("$A:$C").EntireColumn.Delete
Ws.Range("$B:$AC").EntireColumn.Delete
RmvAllColAft Ws, "P"

'== Put Row 1 COl-B 15 columns to Row 2==
Dim Rg As Range, V
Dim J%: For J = 1 To 15
    Set Rg = Ws.Cells(1, J + 1)
    V = Rg.Value
    Set Rg = Ws.Cells(2, J + 1)
    Rg.Value = V
Next
Ws.Range("$1:$1").EntireRow.Delete
Ws.Range("$1:$1").NumberFormat = "MMM YYYY"

'==Vdt Column
Dim O$()
Dim MNy$(): MNy = MonthNy(Y, M, 15)
If Ws.Range("A1").Value <> "Mhd Sku" Then PushS O, "A1 should be [Mhd Sku]"
For J = 0 To 14
    Set Rg = Ws.Cells(1, 2 + J)
    If UCase(Format(Rg.Value, "MMM YYYY")) <> MNy(J) Then PushS O, Rg.Address & " should be [" & MNy(J) & "]"
Next
Wb.Save
If Si(O) = 0 Then
    Wb.Application.WindowState = xlMaximized
    Wb.Save
    PushS O, ""
    PushS O, "Above error is in the working file:"
    PushS O, "Work Folder    : [" & FcWPth & "]"
    PushS O, "Excel Work File: [" & Wb.Name & "]"
    PushS O, ""
    PushS O, "Fix the original worksheet and load again!"
    BrwEr O
End If

'== Ren Field as : Sku M01..M15
Ws.Range("A1").Value = "Sku"
For J = 1 To 15
    Set Rg = Ws.Cells(1, 1 + J)
    Rg.Value = "M" & Format(J, "00")
Next
End Sub
'---===============================================================================================
Private Sub RMhLnk__Tst()
RMhLnk LasMHFc
OpnTbl T_LFc
End Sub
Private Sub RMhLnk(A As StmYM)
CLnkFxw FcWFx(A), MhWsn, T_LFc
End Sub
'--
Private Sub RUdLnk__Tst()
RUdLnk LasUDFc
OpnTbl T_LFc86
OpnTbl T_LFc87
End Sub

Private Sub RUdLnk(A As StmYM)
CLnkFxw FcWFx(A), Wsn86, T_LFc86
CLnkFxw FcWFx(A), Wsn87, T_LFc87
End Sub
'---===============================================================================================
Private Sub RMhImp__Tst(): RMhImp: End Sub
Private Sub RUdImp__Tst(): RUdImp: End Sub
Private Sub RMhImp()
DoCmd.SetWarnings False
RunCQ SqlSelStar_Into_Fm(T_TIFc, T_LFc) & WhFeq("FcTy", FcTy)
RunCQ SqlAddCol(T_TIFc, "Co Byte")
RunCQ SqlUpd(T_TIFc, "Co=86") & WhFeq("CoNm", "HONG KONG DP")
RunCQ SqlUpd(T_TIFc, "Co=87") & WhFeq("CoNm", "MACAO DP")
RunCQ SqlDrpCol_T_F(T_TIFc, "Market,CoNm,FcTy")
End Sub
Private Sub RUdImp()
'Import >IFc86 to #IFc86
'And    >IFc87 to #IFc87
Dim Sel$: Sel = "Select Sku," & ExpandPfxNN("M", 1, 15, "00", Sep:=",")
RunCQ Sel & QpIntoFm(T_TIFc86, T_LFc86)
RunCQ Sel & QpIntoFm(T_TIFc87, T_LFc87)
End Sub
'---===============================================================================================
Private Sub RMhTmpFcSku()
'Inp: >Fc
'Oup: creaet #Fc Co Sku M01..M15
Dim Sel$: Sel = SelSkuM12 & ",Co"
Dim IntoTmpFcSku$: IntoTmpFcSku = QpInto(T_TFc)
Dim Fm$: Fm = QpFm(T_TIFc, "x")
Dim Wh$: Wh = WhM12 & ")"
RunCQ Sel & IntoTmpFcSku & Fm & Wh
RunCQ "alter table [#Fc] add column M13 Double,M14 Double,M15 double"
SetNull
End Sub
Sub RUdTmpFcSku()
'Inp: #IFc86 & #IFc87
'Oup: #Fc
Const Las3M$ = ",Val(Nz(x.M13,0)) As M13," & _
         "Val(Nz(x.M14,0)) As M14," & _
         "Val(Nz(x.M15,0)) As M15"
Dim Sel86$: Sel86 = SelSkuM12 & Las3M & ",CByte(86) as Co"
Dim Sel87$: Sel87 = SelSkuM12 & Las3M & ",CByte(87) as Co"
Dim IntoTmpFc$: IntoTmpFc = QpInto("#Fc")
Dim IntoTmpA$: IntoTmpA = QpInto("#A")
Dim Fm86$: Fm86 = QpFm(T_TIFc86, "x")
Dim Fm87$: Fm87 = QpFm(T_TIFc87, "x")
Dim Wh$: Wh = WhM12 & _
                " or Val(Nz(M13,0))<>0" & _
                " or Val(Nz(M14,0))<>0" & _
                " or Val(Nz(M15,0))<>0)"
Dim InsTmpFc$: InsTmpFc$ = QpInsInto("#Fc")
DoCmd.SetWarnings True
RunCQ Sel86 & IntoTmpFc & Fm86 & Wh
RunCQ InsTmpFc & " " & Sel87 & Fm87 & Wh
DoCmd.SetWarnings False
SetNull
'Set each fields-of-Mnn to null if it is 0
End Sub
'-------------------------------------------------------------------------------------------
Private Sub SetNull()
Dim J%: For J = 1 To 15
    Dim F$: F = "M" & Format(J, "00")
    Dim E$: E = F & "=0"
    Dim S$: S = "Update [#Fc] set " & F & "=Null Where " & E
    RunCQ S
Next
End Sub
Private Function WhM12$()
WhM12 = Wh("CStr(Nz(Sku,''))<>''" & _
        " and (Val(Nz(M01,0))<>0" & _
          " or Val(Nz(M02,0))<>0" & _
          " or Val(Nz(M03,0))<>0" & _
          " or Val(Nz(M04,0))<>0" & _
          " or Val(Nz(M05,0))<>0" & _
          " or Val(Nz(M06,0))<>0" & _
          " or Val(Nz(M07,0))<>0" & _
          " or Val(Nz(M08,0))<>0" & _
          " or Val(Nz(M09,0))<>0" & _
          " or Val(Nz(M10,0))<>0" & _
          " or Val(Nz(M11,0))<>0" & _
          " or Val(Nz(M12,0))<>0")
End Function
Private Function SelSkuM12$(): SelSkuM12 = "Select CStr(x.Sku) As Sku," & _
    "Val(Nz(x.M01,0)) As M01," & _
    "Val(Nz(x.M02,0)) As M02," & _
    "Val(Nz(x.M03,0)) As M03," & _
    "Val(Nz(x.M04,0)) As M04," & _
    "Val(Nz(x.M05,0)) As M05," & _
    "Val(Nz(x.M06,0)) As M06," & _
    "Val(Nz(x.M07,0)) As M07," & _
    "Val(Nz(x.M08,0)) As M08," & _
    "Val(Nz(x.M09,0)) As M09," & _
    "Val(Nz(x.M10,0)) As M10," & _
    "Val(Nz(x.M11,0)) As M11," & _
    "Val(Nz(x.M12,0)) As M12"
End Function

'---========================================================================================
Private Sub RChk__Tst():
Call RMhTmpFcSku: RMhChk
Call RUdTmpFcSku: RUdChk
End Sub
Private Sub RMhChk()
'Inp: #FcSku
'BrwEr
'     1 Vdt Dup-Sku in #FcSku
'     2 Sku is belong to correct Stm
'     Thw Er and Show the NotePad Message
BrwEr RChk(T_TFc, "M")
End Sub
Private Sub RUdChk()
'Inp: #IFc86 & #IFc87
'BrwEr
'     1 Vdt Dup-Sku in #FcSku
'     2 Sku is belong to correct Stm
'     Thw Er and Show the NotePad Message
BrwEr RChk(T_TFc, "U")
End Sub

Private Function RChk$(T$, Stm$)
RChk = AddLinesAp(VdtDupSku(T), VdtWrongStm(T, Stm), VdtSku(T))
End Function

Private Function VdtSku$(T$)
VdtSku = VdtSkuy(SyzCQ(SqlSel_F_T("Sku", T, Dis:=True)))
End Function

Function VdtSkuy$(Skuy$())
Dim Good$(): Good = SyzCQ("Select Sku from Sku")
Dim Bad$(): Bad = MinusSy(Skuy, Good)
If Si(Bad) > 0 Then
    Const L1$ = "Following Sku is not found in tha Sku table" & vbCrLf
    VdtSkuy = L1 & vbCrLf & JnCrLf(AmAddPfxTab(Bad)) & vbCrLf
End If
End Function

Private Function VdtDupSku$(T)
Dim SKU$(): SKU = SyzCQ(QpSelDist("Sku", T_TFc))
VdtDupSku = DupEleMsgl(SKU, "Sku")
End Function

Private Function VdtWrongStm$(T$, Stm$)
Dim SKU$(): SKU = SyzCQ(QpSelDist("Sku", T))
Dim StmSku$(): StmSku = SkuyzStm(Stm)
Dim O$()
Dim I: For Each I In SKU
    If Not HasEle(StmSku, I) Then
        PushI O, I
    End If
Next
Dim Lines$: Lines = FmtQQ("Following Sku is not [?]-Sku", StreamzStm(Stm)) & vbCrLf
If Si(O) > 0 Then VdtWrongStm = Lines & TabAy(O) & vbCrLf
End Function
Function TabAy$(Ay)
Dim O$():
Dim I: For Each I In Itr(Ay)
    PushS O, vbTab & I
Next
TabAy = Join(O, vbCrLf)
End Function
'---=================================================================================================== Load
Private Sub UpdTbPHStkDays7__Tst()
Dim A As StmYM
A = LasUDFc
RMhLnk A
RMhImp
RMhTmpFcSku
UpdTbPHStkDays7 A
'A = LasMhFc
'RUdLnk A
'RUdImp
RUdTmpFcSku
UpdTbPHStkDays7 A
End Sub
Private Sub UpdTbPHStkDays7(A As StmYM)
'Inp: #FcSku =              Co Sku M01..15 Assume all Sku is in the Stm as @A
'Oup: FcSku   = VerYY VerMM Co Stm SKu M01..15 So #A is just add
'Stp: FcSku   : Delete where VerYY VerMM Co Stm
'     FcSku   : Append from #A

'-- #Sku -> Delete FcSku
RunCQ "Delete * from FcSku" & WhereFcStm(A)
With A
Dim Sql$: Sql = FmtQQ("insert into FcSku" & _
    "(VerYY,VerMM,Co,  Stm,Sku,M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,M13,M14,M15) Select" & _
    " ?    ,    ?,Co,'?',Sku,M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,M13,M14,M15" & _
    " from [#Fc]", .Y, .M, .Stm)
    End With
DoCmd.SetWarnings True
RunCQ Sql
DoCmd.SetWarnings False
End Sub
'---===============================================================================================
Private Sub RMhDrpTmp()
DrpCTbAp T_LFc, T_TIFc, T_TFc
End Sub
Private Sub RUdDrpTmp()
DrpCTbAp T_LFc86, T_LFc87, T_TIFc86, T_TIFc87, T_TFc
End Sub
Function FcWFx$(A As StmYM)
FcWFx = FcWPth & "Wrk " & FcIFxFn(A)
End Function

Private Function IFx$(A As StmYM)
IFx = FcIFx(A)
End Function

Function FcWPth$()
Static P$
If P = "" Then
    P = IPth & "Wrk\"
    If Dir(P, vbDirectory) = "" Then MkDir P
End If
FcWPth = P
End Function

Private Function IPth$()
Static P$
If P = "" Then
    P = AppIPth & "Forecast\"
    If Dir(P, vbDirectory) = "" Then MkDir P
End If
IPth = P
End Function

'---=============================================================================================== Las
Function LasFc() As YM
LasFc = YMzYYMM(VzCQ("select Max(VerYY*100+VerMM) from Fc"))
End Function
Function LasStmFc(Stm$) As YM
LasStmFc = YMzYYMM(VzCQ("select Max(VerYY*100+VerMM) from Fc" & StmBexp(Stm)))
End Function
Function LasUDFc() As StmYM
LasUDFc = StmYM1("U", LasStmFc("U"))
End Function
Function LasMHFc() As StmYM
LasMHFc = StmYM1("M", LasStmFc("M"))
End Function
'---==============================================================================================
Private Function UdWsn$(Co As Byte)
Select Case Co
Case 87: UdWsn = Wsn87
Case 86: UdWsn = Wsn86
End Select
End Function
