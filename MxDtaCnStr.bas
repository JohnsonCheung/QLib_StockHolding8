Attribute VB_Name = "MxDtaCnStr"
Option Compare Text
Option Explicit
Const CNs$ = "sdfsdf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDtaCnStr."
Function CnStrzDbt$(D As Database, T): CnStrzDbt = DtaSrczScvl(D.TableDefs(T).Connect): End Function


Function DaoCnStrzFb$(Fb)
DaoCnStrzFb = ";DATABASE=" & Fb & ";"
End Function

Function OleCnStrzFb$(Fb) 'Return a connection used as WbConnection
OleCnStrzFb = "OLEDb;" & AdoCnStrzFb(Fb)
End Function

Function WcCnStrzFb$(Fb)
WcCnStrzFb = OleCnStrzFb(Fb)
'WcCnStrzFb = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
'WcCnStrzFb = FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
End Function

Function DaoCnStrzFx$(Fx)
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
'Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=C:\Users\sium\Desktop\TaxRate\sales text.xlsx;TABLE=Sheet1$
Dim O$
Select Case LCase(Ext(Fx))
Case ".xlsx":: O = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & Fx & ";"
Case ".xls": O = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & Fx & ";"
Case Else: Stop
End Select
DaoCnStrzFx = O
End Function

Function CnStrAy(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushNB CnStrAy, CnStrzT(D, T)
Next
End Function

Function DaoCnStrzFcsv$(Fcsv)
Dim Fn$: Fn = RmvExt(Fcsv) & "#Csv"
DaoCnStrzFcsv = FmtQQ("Text;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=?;TABLE=?", Pth(Fcsv), Fn)
''Text;DSN=Delta_Tbl_08052203_20080522_033948 Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=C:\Tmp;TABLE=Delta_Tbl_08052203_20080522_033948#csv

End Function
