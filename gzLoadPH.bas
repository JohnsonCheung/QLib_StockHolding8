Attribute VB_Name = "gzLoadPH"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzLoadPH."
Option Base 0

Sub LoadPH(Optional Frm As Access.Form)
Dim IFx$: IFx = PHIFx
Const Wsn$ = "Sheet1"
Const FldNmCsv = "Product hierarchy,Level no#,Description"
Const XlsTyCsv = " T                ,T        ,T"
ChkWsCol IFx, "Sheet1", FldNmCsv, XlsTyCsv
CLnkFxw IFx, Wsn, ">PH"
DoCmd.SetWarnings False
'-------------------
Sts "Running import query ....."
'Crt #PH = PH Lvl Des
RunCQ "Select [Product hierarchy] as PH,[Level no#] as Lvl,Description as Des into [#PH] from [>PH]"

'Dlt #PH ! For
DltTmpPH
    
'Ins ProdHierarchy
'Upd ProdHierarchy
'Drp #PH
DoCmd.SetWarnings True
RunCQ "Insert into ProdHierarchy Select x.PH,CByte(Val(x.Lvl)) as Lvl,x.Des from [#PH] x left Join ProdHierarchy a on x.PH=a.PH where a.PH is null"
RunCQ "Update ProdHierarchy x inner join [#PH] a on x.PH=a.PH set x.Des=a.Des, x.Lvl=a.Lvl, DteUpd=Now where x.Des<>a.Des or x.Lvl<>CByte(Val(a.Lvl))"

'Upd ProdHierarchy->WithOH
'Upd ProdHierarchy->Sno
'Upd ProdHierarchy->Srt
RfhTbPH_Fld_WithOHxxx
RfhTbPH_FldSno
RfhTbPH_FldSrt

'Drp
DrpCTT ">PH #PH"
'==
If Not IsNothing(Frm) Then Frm.Requery
End Sub

Private Sub DltTmpPH()
Dim N%

N = CurrentDb.OpenRecordset("Select Count(*) from [#PH] where Trim(Nz(PH,''))=''").Fields(0).Value
If N > 0 Then
    MsgBox "There are [" & N & "] records with [Product Hierarchy] is blank.  They are omitted!", vbInformation
    RunCQ "Delete * from [#PH] where Trim(Nz(PH,''))='')"
End If
    
N = CurrentDb.OpenRecordset("Select Count(*) from [#PH] where Len(Nz(PH,''))>10").Fields(0).Value
If N > 0 Then
    MsgBox "There are [" & N & "] records with [Product Hierarchy] length is >10.  They are omitted!", vbInformation
    RunCQ "Delete * from [#PH] where Len(Nz(PH,''))>10"
End If
    
N = CurrentDb.OpenRecordset("Select Count(*) from [#PH] where Not (CLng(Val(Lvl)) between 1 and 5)").Fields(0).Value
If N > 0 Then
    MsgBox "There are [" & N & "] records with Level# not between 1 and 5.  They are omitted!", vbInformation
    RunCQ "Delete * from [#PH] where Not (CLng(Val(Lvl)) between 1 and 5)"
End If
    
N = CurrentDb.OpenRecordset("Select Count(*) from [#PH] where trim(Nz(Des,''))=''").Fields(0).Value
If N > 0 Then
    MsgBox "There are [" & N & "] records with Blank Description.  They are omitted!", vbInformation
    RunCQ "Delete * from [#PH] where trim(Nz(Des,'')=''"
End If

End Sub

Function PHIFx$()
PHIFx = CPmv("PH_InpFx")
End Function
