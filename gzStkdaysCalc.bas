Attribute VB_Name = "gzStkdaysCalc"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzStkdaysCalc."
Private Type StkDaysRem
    StkDays As Integer
    RemSC As Double  ' See Subr-StkDaysRem
End Type

'Sub T_OH(): ROH SampYmd: End Sub
'Sub T_Fc(): RFc SampYmd: End Sub
'Sub T_Days(): TmpDays7 SampYmd: End Sub
'Sub T_Load(): UpdTbPHStkDays7 SampYmd: End Sub
'Sub T_Calc(): UpdTbPHStkDays7_FldStkDays_andRemSc SampYmd: End Sub
Private Function SampYmd() As Ymd: SampYmd = Ymd(19, 12, 24): End Function

Sub UpdTbPHStkDays7_FldStkDays_andRemSc__Tst()
UpdTbPHStkDays7_FldStkDays_andRemSc LasOHYmd
End Sub

Sub UpdTbPHStkDays7_FldStkDays_andRemSc(A As Ymd) 'Upd Tb-PHStkDays{7}->StkDays/RemSc
'Inp: OH
'Inp: FcSku = VerYY VerMM Co Stm Sku M01..15
'Where {7} of PHStkDays{7} = {Sku Stm Bus L1..4}
'---
Sts "Update Stock Days ....":
TmpScOH7_ByYmd A        '$ScOH{7} are created        = Co Stm {7} SC
TmpFc7 A       '$Fc{7} are created        = Co Stm {7} M01..M15
TmpDays7 A     '$Days{7} are created      = Co Stm {7} StkDays
UpdTbPHStkDays7 A     'Load $Days{7} into PHStkDays{7} = Y M D Co Stm {7} StkDays
DrpCPH7Tbl "$Fc?"
DrpCPH7Tbl "$ScOH?"
DrpCPH7Tbl "$Days?"
End Sub

Private Sub TmpDays7(A As Ymd)
'Inp: $ScOH{7} = Co Stm {7} SC
'Inp: $Fc{7} = Co Stm {7} M01..15
'Oup: $Days{7} = Co Stm {7} StkDays
'---
'#1 Tmp: #Fc{7}OH = Co Stm {7} M01..15 SC           ! SC is added at end of $Fc{7} to become #Fc{7}X
'#2 Tmp: $Fc{7}OH = ..                    StkDays   ! Add StkDays.  Each Record has enought data to calc StkDays
'#3 Oup: $Days{7} = Co Stm {7} StkDays              ! From $Fc{7}OH, just remove the M01..M15
'-----
'== Stp-TmpFcXOH: $Fc{7}OH
'      Fm : $Fc{7}
Const Sel$ = "Select x.*," & _
"a.M01,a.M02,a.M03," & _
"a.M04,a.M05,a.M06," & _
"a.M07,a.M08,a.M09," & _
"a.M10,a.M11,a.M12," & _
"a.M13,a.M14,a.M15,CInt(0) as StkDays,CDbl(0) as RemSC"
RunCQ Sel & " into [#FcSkuOH] from [$ScOHSku] x left join [$FcSku] a on a.Co=x.Co and a.Stm=x.Stm and a.Sku=x.Sku"
RunCQ Sel & " into [#FcBusOH] from [$ScOHBus] x left join [$FcBus] a on a.Co=x.Co and a.Stm=x.Stm and a.BusArea=x.BusArea"
RunCQ Sel & " into [#FcStmOH] from [$ScOHStm] x left join [$FcStm] a on a.Co=x.Co and a.Stm=x.Stm"
RunCQ Sel & " into [#FcL1OH]  from [$ScOHL1]  x left join [$FcL1]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL1=x.PHL1"
RunCQ Sel & " into [#FcL2OH]  from [$ScOHL2]  x left join [$FcL2]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL2=x.PHL2"
RunCQ Sel & " into [#FcL3OH]  from [$ScOHL3]  x left join [$FcL3]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL3=x.PHL3"
RunCQ Sel & " into [#FcL4OH]  from [$ScOHL4]  x left join [$FcL4]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL4=x.PHL4"
'
'== Stp-UpdStkDays $Fc{7}OH: Update StkDays & RemSC
UpdTblT_FldStkDaysRem "#FcSkuOH", A
UpdTblT_FldStkDaysRem "#FcStmOH", A
UpdTblT_FldStkDaysRem "#FcL1OH", A
UpdTblT_FldStkDaysRem "#FcL2OH", A
UpdTblT_FldStkDaysRem "#FcL3OH", A
UpdTblT_FldStkDaysRem "#FcL4OH", A
UpdTblT_FldStkDaysRem "#FcBusOH", A
'== Stp-Days Create: $Days{7} from #Fc{7}OH just drop all the M01..15
BB_CrtTmpDaysAndDrpM1To15 "Sku"  '$DaysSku is created
BB_CrtTmpDaysAndDrpM1To15 "Stm"
BB_CrtTmpDaysAndDrpM1To15 "Bus"
BB_CrtTmpDaysAndDrpM1To15 "L1"
BB_CrtTmpDaysAndDrpM1To15 "L2"
BB_CrtTmpDaysAndDrpM1To15 "L3"
BB_CrtTmpDaysAndDrpM1To15 "L4"

'---=
RunCQ "Drop Table [#FcSkuOH]"
RunCQ "Drop Table [#FcStmOH]"
RunCQ "Drop Table [#FcBusOH]"
RunCQ "Drop Table [#FcL1OH]"
RunCQ "Drop Table [#FcL2OH]"
RunCQ "Drop Table [#FcL3OH]"
RunCQ "Drop Table [#FcL4OH]"
End Sub
Private Sub BB_CrtTmpDaysAndDrpM1To15(LvlItm$)
RunCQ "Select * into [$Days" & LvlItm & "] from [#Fc" & LvlItm & "OH]"
RunCQ "Alter Table [$Days" & LvlItm & "] drop column " & _
"M01,M06,M11," & _
"M02,M07,M12," & _
"M03,M08,M13," & _
"M04,M09,M14," & _
"M05,M10,M15"
End Sub
Private Sub FcAy__Tst()
Dim Rs As DAO.Recordset: Set Rs = CurrentDb.TableDefs("FcSku").OpenRecordset
With Rs
    Dim F!: F = RemDaysFactor(Now)
    Dim Fc#()
    While Not .EOF
        Fc = FcAy(Rs, F)
        Dim J%: For J = 0 To Si(Fc) - 1
            Debug.Print Fc(J);
        Next
        Debug.Print
        .MoveNext
    Wend
End With
End Sub

Private Function FcAy(Rs As DAO.Recordset, FstMthRemDaysFac!) As Double()
'@Rs: It has M01..15
'Return : the FcAy with FstMth adjust to remaining days as in @FstMthRemDaysFac
'         Trim all the end element if it is zero
Dim OFc#(): ReDim OFc(14)
Dim J%: For J = 0 To 14
    OFc(J) = Nz(Rs.Fields("M" & Format(J + 1, "00")).Value, 0)
Next
OFc(0) = OFc(0) * FstMthRemDaysFac ' Adjust the first Month
For J = 14 To 0 Step -1
    If OFc(J) <> 0 Then
        ReDim Preserve OFc(J)
        FcAy = OFc
        Exit Function
    End If
Next
'-- All Fc is Zero, just return
End Function

Private Sub Dayy__Tst()
DmpAy Dayy(Ymd(19, 11, 28))
End Sub

Private Function Dayy(A As Ymd) As Byte()
'Return : 15 month's days with @@Days(0) is the remaining date of the month of @A
Dim O() As Byte: ReDim O(14)
Dim D As Date: D = DtezYmd(A)
O(0) = RemDays(D)
Dim J%: For J = 1 To 14
    D = FstDteNxtMth(D)
    O(J) = NDay(D)
Next
Dayy = O
End Function

Private Sub UpdTblT_FldStkDaysRem(T$, A As Ymd)
'@T: :#Fc{}OH: StkDays RemSC SC M01..15
'Oup: @T->StkDays & RemSC are updated
Dim Y As Byte, M As Byte: Y = A.Y: M = A.M
Dim F!:    F = RemDaysFactor(DtezYmd(A))
Dim D() As Byte: D = Dayy(A)
Dim Rs As DAO.Recordset: Set Rs = RszCT(T)
With Rs
    Dim Fc#(), SC#
    While Not .EOF
        Fc = FcAy(Rs, F)
        SC = !SC
        Dim C As StkDaysRem: C = StkDaysRem(SC, Fc, D)
        .Edit
            !StkDays = C.StkDays
            !RemSC = C.RemSC
        .Update
        .MoveNext
    Wend
End With
End Sub

Private Sub TmpFc7(A As Ymd)
TmpFc_ByYM YM(A.Y, A.M)
End Sub

Private Sub UpdTbPHStkDays7(A As Ymd)
'Inp: #StkDays{7}  =          Stm {7} StkDays
'Oup: PHStkDays{7} = Y M D Co Stm {7} StkDays, {7} = Sku Stm Bus L1..4
DoCmd.SetWarnings False
Dim W$: W = OHYmdBexp(A)
RunCQ "Delete * from PHStkDaysSku" & W
RunCQ "Delete * from PHStkDaysBus" & W
RunCQ "Delete * from PHStkDaysStm" & W
RunCQ "Delete * from PHStkDaysL1" & W
RunCQ "Delete * from PHStkDaysL2" & W
RunCQ "Delete * from PHStkDaysL3" & W
RunCQ "Delete * from PHStkDaysL4" & W
Dim Sql$, Y As Byte, M As Byte, D As Byte
With A
    Y = .Y
    M = .M
    D = .D
End With

'Sku
Sql = FmtQQ("Insert into PHStkDaysSku" & _
" (YY, MM, DD,Co,Stm,Sku,StkDays,RemSC) Select" & _
" ?,?,?,Co,Stm,Sku,StkDays,RemSC from [$DaysSku]", Y, M, D)
RunCQ Sql

'Stm
Sql = FmtQQ("Insert into PHStkDaysStm" & _
" (YY, MM, DD,Co,Stm,StkDays,RemSC) Select" & _
" ?,?,?,Co,Stm,StkDays,RemSC from [$DaysStm]", Y, M, D)
RunCQ Sql

'Bus
Sql = FmtQQ("Insert into PHStkDaysBus" & _
" (YY, MM, DD,Co,Stm,BusArea,StkDays,RemSC) Select" & _
" ?,?,?,Co,Stm,BusArea,StkDays,RemSC from [$DaysBus]", Y, M, D)
RunCQ Sql

'L1
Sql = FmtQQ("Insert into PHStkDaysL1" & _
" (YY, MM, DD,Co,Stm,PHL1,StkDays,RemSC) Select" & _
" ?,?,?,Co,Stm,PHL1,StkDays,RemSC from [$DaysL1]", Y, M, D)
RunCQ Sql

'L2
Sql = FmtQQ("Insert into PHStkDaysL2" & _
" (YY, MM, DD,Co,Stm,PHL2,StkDays,RemSC) Select" & _
" ?,?,?,Co,Stm,PHL2,StkDays,RemSC from [$DaysL2]", Y, M, D)
RunCQ Sql

'L3
Sql = FmtQQ("Insert into PHStkDaysL3" & _
" (YY, MM, DD,Co,Stm,PHL3,StkDays,RemSC) Select" & _
" ?,?,?,Co,Stm,PHL3,StkDays,RemSC from [$DaysL3]", Y, M, D)
RunCQ Sql

'L4
Sql = FmtQQ("Insert into PHStkDaysL4" & _
" (YY, MM, DD,Co,Stm,PHL4,StkDays,RemSC) Select" & _
" ?,?,?,Co,Stm,PHL4,StkDays,RemSC from [$DaysL4]", Y, M, D)
RunCQ Sql
End Sub

Private Function StkDaysRem(SC#, Fc#(), Days() As Byte) As StkDaysRem
'@Fc :SCAy #Forecast-of-Each-Month-in-StdCase#
'@SC :SC   #OH-in-SC#
'@Days     #Days-of-each-Month# ! @Days(0) has adjusted according the Given Ymd
':StkDays: :Days ! Number of days can cover the Fc-Quantity for each months.
':RemSC:   :Dbl  ! #Remaing-SC#
'                ! It may in 1 of 3 conditions:
'               !   1. If All @SC can be consumed:                    RemSc <= 0
'               !   2. No Forecast, that means zero element in @Fc:   RemSc <= -1
'               !   3. {Days}+     (THe @SC is not                    RemSc <= A positive number
'Return :StkDays !  {SC#} (StdCase) by {Fc#()} and {Days}
If Si(Fc) = 0 Then
    StkDaysRem.StkDays = 9999
    StkDaysRem.RemSC = SC
    Exit Function
End If
Dim RemSC#: RemSC = SC
Dim ODays#, iSC#
Dim J%: For J = 0 To UBound(Fc)
    If RemSC < Fc(J) Then StkDaysRem.StkDays = Round(ODays + Days(J) * RemSC / Fc(J)): Exit Function
    RemSC = RemSC - Fc(J)
    ODays = ODays + Days(J)
Next
If RemSC > 0 Then
    StkDaysRem.StkDays = Round(ODays)
    StkDaysRem.RemSC = RemSC
    Exit Function
End If
StkDaysRem.StkDays = Round(ODays)
End Function
