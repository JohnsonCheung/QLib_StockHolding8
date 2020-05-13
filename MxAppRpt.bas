Attribute VB_Name = "MxAppRpt"
Option Compare Text
Option Explicit
Const CNs$ = "App"
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxAppRpt."
Type TFxww: Fx As String: Wny() As String: End Type
Type TFbtt: Fb As String: Tny() As String: End Type
Type Inp: Fx As TFxww: Fb As TFbtt: End Type
Type RptPm
    Nm As String
    Inp As Inp
    LnkImpSrc() As String
    WPth As String
    WFb As String
    TpFx As String
    OuPfxy() As String
    GenOupFun As String
    FmtWbFun As String
End Type

Sub Rpt(P As RptPm) 'Gen&Vis OupFx using LidPm as NxtFfnzAva.
With P
':                              CpyFfnAyzIfDif .InpFilSrc, .WPth         ' <== Cpy inp fil to wpth
:                              EnsFb .WFb                              ' <== Crt wrk fb
Dim W As Database:     Set W = Db(.WFb)
':                              LnkImp Sy(.InpFilSrc, .LnkImpSrc), W    ' <== LnkImp
:                              Run .GenOupFun           ' <== Gen oup tbl
:                              W.Close
':                              CpyFfn .TpFx, .OupFx                    ' <== Cpy to OupFx.  Assume OupFx is always new
':                              RfhFx .OupFx, .WFb
Dim Wb1 As Excel.Workbook:  '   Set Wb1 = WbzFx(.OupFx)
':                              CpyInp P, Wb1                          ' <== Cpy inp ws
':                              Run .FunoFmtWb, Wb1                       ' <== Fmt wb
:                              Wb1.Save                                 ' <== Sav
':                              If Not .IsOpn Then Wb1.Close            ' <== KeepOpn?
End With
End Sub

Private Sub Rpt__Tst()
Dim WDb As Database
GoSub Z
Exit Sub
Z:
    Set WDb = Nothing
    GoTo Tst
Tst:
    'Rpt_Oup WDb, B
    Return
End Sub
