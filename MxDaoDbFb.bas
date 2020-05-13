Attribute VB_Name = "MxDaoDbFb"
Option Compare Text
Option Explicit
Const CNs$ = "sw"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbFb."

Function ArszFbq(Fb, Q) As ADODB.Recordset
Set ArszFbq = CnzFb(Fb).Execute(Q)
End Function

Sub ArunFbq(Fb, Q) ' Use AdoDb to @Q in @Fb
CnzFb(Fb).Execute Q
End Sub

Sub AsgFbtStr(FbtStr$, OFb$, OT$)
If FbtStr = "" Then
    OFb = ""
    OT = ""
    Exit Sub
End If
AsgBrk OFb, OT, _
    FbtStr, "].["
If FstChr(OFb) <> "[" Then Stop
If LasChr(OT) <> "]" Then Stop
OFb = RmvFstChr(OFb)
OT = RmvLasChr(OT)
End Sub

Function CrtFb(Fb) As Database ' Crt Fb and return as ::Db
Set CrtFb = DAO.DBEngine.CreateDatabase(Fb, dbLangGeneral)
End Function

Sub EnsFb(Fb)
If NoFfn(Fb) Then CrtFb Fb
End Sub

Function DbzFb(Fb) As Database
Set DbzFb = Db(Fb)
End Function

Function Db(Fb) As Database
Set Db = DAO.DBEngine.OpenDatabase(Fb)
End Function

'--
Sub DrpFbt(Fb, T)
CatzFb(Fb).Tables.Delete T
End Sub

Function DrszFbq(Fb, Q) As Drs
DrszFbq = DrszRs(Rs(Db(Fb), Q))
End Function

Function DrszQ(D As Database, Q) As Drs
DrszQ = DrszRs(Rs(D, Q))
End Function

Private Sub OupTnyzFb__Tst()
D OupTnyzFb(CFb)
End Sub

Function OupTnyzFb(Fb) As String()
OupTnyzFb = OupTny(Db(Fb))
End Function

Function WszFbq(Fb, Q, Optional Wsn$) As Worksheet
Set WszFbq = WszDrs(DrszFbq(Fb, Q), Wsn:=Wsn)
End Function

Private Sub BrwFb__Tst()
BrwFb DutyDtaFb
End Sub

Private Sub HasFbt__Tst()
Ass HasFbt(DutyDtaFb, "SkuB")
End Sub


Private Sub TnyzFb__Tst()
DmpAy TnyzFb(DutyDtaFb)
End Sub

Private Sub WszFbq__Tst()
VisWs WszFbq(DutyDtaFb, "Select * from KE24")
End Sub
