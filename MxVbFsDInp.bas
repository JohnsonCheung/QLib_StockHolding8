Attribute VB_Name = "MxVbFsDInp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsDInp."

Function SampDoInp() As Drs
Erase XX
X "MB52 C:\Users\user\Desktop\Mhd\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
X "UOM  C:\Users\user\Desktop\Mhd\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
X "ZHT1 C:\Users\user\Desktop\Mhd\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
SampDoInp = DrszTRst("FilKd Ffn", XX)
End Function

Function EoMsgDrs(Msg$, A As Drs) As String()
If NoReczDrs(A) Then Exit Function
Erase XX
XLin Msg
XDrs A
XLin
EoMsgDrs = XX
End Function

Private Sub EoDInp__Tst()
Brw EoDInp(SampDoInp)
End Sub

Function EoDInp(DInp As Drs) As String()
'@DInp
Dim E1$(), E2$(), E3$()
E1 = EoDupCol(DInp, "Ffn")
E2 = EoDupCol(DInp, "FilKd")
E3 = EoFfnMiszD(DInp)
EoDInp = Sy(E1, E2, E3)
End Function

Function EoFfnMiszD(WiFfn As Drs) As String()
Dim I%: I = IxzAy(WiFfn.Fny, "Ffn")
Dim Dr, Dy(): For Each Dr In Itr(WiFfn.Dy)
    If NoFfn(Dr(I)) Then PushI Dy, Dr
Next
Dim B As Drs: B = Drs(WiFfn.Fny, Dy)
EoFfnMiszD = EoMsgDrs("File not exist", B)
End Function
