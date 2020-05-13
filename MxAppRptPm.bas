Attribute VB_Name = "MxAppRptPm"
Option Explicit
Option Compare Text
Const CNs$ = "Rpt.Pm"
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxAppRptPm."
Type Oup
    A As String
End Type
Type Inp
    A As String
End Type
Type RptPm
    Inp As Inp
    Oup As Oup
End Type

Function RptPmTp$()
#If False Then
Fx:  Fxn Ffn
    MB52 C
    MXAA C:\AsWW
Ws: Fbn Wsnn
    MB52 Sheet Sheet2
    MXAA Sheet1 SHeet2
Fb: Nm Ffn
    DDD C:\sdfsdf
    FF  C:\sdfdf
Fbt: Fbn Tbnn
    DDD AA BB
    CCC BB DD
FbTbl:
    AAA MB52 Sheet1
    BBB MB52 Sheet1
FxTbl
    DDD AAA
OupFx
    MB32 C:\LJKLKJDf
    
OupWs
    
OupLo
    
OupPt

#End If

Const A_1$ = "E Mem | Mem Req AlZZLen" & _
vbCrLf & "E Txt | Txt Req" & _
vbCrLf & "E Crt | Dte Req Dft=Now" & _
vbCrLf & "E Dte | Dte" & _
vbCrLf & "F Amt * | *Amt" & _
vbCrLf & "F Crt * | CrtDte" & _
vbCrLf & "F Dte * | *Dte" & _
vbCrLf & "F Txt * | Fun * Txt" & _
vbCrLf & "F Mem * | Lines" & _
vbCrLf & "T Sess | * CrtDte" & _
vbCrLf & "T Msg  | * Fun *Txt | CrtDte" & _
vbCrLf & "T Lg   | * Sess Msg CrtDte" & _
vbCrLf & "T LgV  | * Lg Lines" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt" & _
vbCrLf & "D . Msg | ..."
'LnkSpecTp = A_1
End Function
