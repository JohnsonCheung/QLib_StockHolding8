Attribute VB_Name = "gzRfhTbSku_Ovr"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "gzRfhTbSku_Ovr."
Sub RfhTbSku_Ovr()
'BusArea
'BusAreaSap
'BusAreaOvr
'Litre/BtlSap
'Litre/BtlOvr
'Litre/Btl
RunCQ "Update Sku Set BusArea='' Where BusArea Is Null"
RunCQ "Update Sku Set BusAreaOvr='' Where BusAreaOvr Is Null"
RunCQ "Update Sku Set BusAreaSap='' Where BusAreaSap Is Null"
RunCQ "Update Sku Set [Litre/Btl]=0 Where [Litre/Btl] is null"
RunCQ "Update Sku Set [Litre/BtlSap]=0 Where [Litre/BtlSap] is null"
RunCQ "Update Sku Set [Litre/BtlOvr]=0 Where [Litre/BtlOvr] is null"

RunCQ "Update Sku set BusArea=Trim(IIf(BusAreaOvr='',BusAreaSap,BusAreaOvr)) where BusArea<>Trim(IIf(BusAreaOvr='',BusAreaSap,BusAreaOvr))"
RunCQ "Update Sku set [Litre/Btl]=IIf([Litre/BtlOvr]=0,[Litre/BtlSap],[Litre/BtlOvr]) where [Litre/Btl]<>IIf([Litre/BtlOvr]=0,[Litre/BtlSap],[Litre/BtlOvr])"
End Sub
