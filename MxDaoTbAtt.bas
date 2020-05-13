Attribute VB_Name = "MxDaoTbAtt"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CNs$ = "Attn"
Const CMod$ = CLib & "MxDaoTbAtt."
Type Attd
    Db As Database
    RecRs As DAO.Recordset '..Attn..
    FldRs As DAO.Recordset2 '.
End Type
Public Const AttFF$ = "AttId Attn Att"
Public Const AttdFF$ = "AttdId AttId Fn FilTim FilSi"

Function FstAttFn$(D As Database, Attn$) 'Ret : fst attachment fn in the att fld of att tbl, if no fn, return blnk @@
Const CSub$ = CMod & "AttFn"
Dim A As Attd: A = Attd(D, Attn) ' if @Attn in exist in Tbl-Attn, a rec will created
Dim F As Recordset: Set F = A.FldRs
If NoRec(F) Then Thw CSub, "[Attn] has no attachment files", "Attn", Attn
F.MoveFirst
FstAttFn = F!FileName
End Function

Function AttTim(D As Database, Attn, Attf) As Date: AttTim = VzQQ(D, W1Sql(D, Attn, Attf)):                       End Function
Function W1Sql$(D As Database, Attn, Attf):          W1Sql = FmtQQ(W1Tp, AttId(D, Attn), Attf):                   End Function
Function W1Tp$():                                    W1Tp = "Select FilTim from Attd where AttId=? and Attf='?'": End Function

Function AttFnAy(A As Attd) As String():                  AttFnAy = SyzRs(A.FldRs, "FileName"):                                 End Function
Function AttFnAyzNm(D As Database, Attn$) As String(): AttFnAyzNm = AttFnAy(Attd(D, Attn)):                                     End Function
Function AttSi&(D As Database, Attn$):                      AttSi = VzSskv(D, "Attn", "FilSz", Attn):                        End Function
Function AttId&(D As Database, Attn):                       AttId = VzQQ(D, "Select AttId from Attd where Attn='?'", Attn):     End Function
Sub DltAtt(D As Database, Attn$):                                   D.Execute FmtQQ("Delete * from Attn where Attn='?'", Attn): End Sub
Function TbAttDrs(D As Database) As Drs:                 TbAttDrs = DrszT(D, "Attn"):                                           End Function
Function FnyzAttFld(D As Database) As String():        FnyzAttFld = FnyzRs(Attd(D, "*Dft").FldRs):                              End Function
Function FnyzCAttFld() As String():                   FnyzCAttFld = FnyzAttFld(CDb):                                            End Function

Function IsAttOlderzD(D As Database, Attn$, Ffn$) As Boolean
Const CSub$ = CMod & "IsAttOlder"
Dim ATim$:   ATim = AttTim(D, Attn, Fn(Ffn))
Dim FTim$:   FTim = DtezFfn(Ffn)
Dim AttIs$: AttIs = IIf(ATim > FTim, "new", "old")
Dim M$:         M = "Attn is " & AttIs
Inf CSub, M, "Attn Ffn AttTim FfnTim AttIs-Old-or-New?", Attn, Ffn, ATim, FTim, AttIs
End Function

Function IsAttOneFil(D As Database, Attn$) As Boolean
Debug.Print "DbAttHasOnlyFile: " & Attd(D, Attn).FldRs.RecordCount
IsAttOneFil = Attd(D, Attn).FldRs.RecordCount = 1
End Function

Function CNAtt%():                     CNAtt = NAttzD(CDb):           End Function
Function NAtt%(D As Database, Attn):   NAtt = W2NAtt(Attd(D, Attn)): End Function
Function NAttzD%(D As Database)
Dim Attn: For Each Attn In Itr(AttNy(D))
    NAttzD = NAttzD + NAtt(D, Attn)
Next
End Function
Private Function W2NAtt%(D As Attd):  W2NAtt = NReczRs(D.FldRs):      End Function


Function AttnzAttd$(A As Attd): AttnzAttd = A.RecRs!Attn: End Function
Function Attn$(A As Attd): Attn = A.RecRs!Attn: End Function
Function CAttd(Attn$) As Attd: CAttd = Attd(CDb, Attn): End Function

Function TmpAttDb() As Database
'Ret: a tmp db with tb-Att&Attd @@
Dim O As Database: Set O = TmpDb
EnsTbAtt O
Set TmpAttDb = O
End Function

Function Attd(D As Database, Attn) As Attd
'Ret: :Attd ! which keeps :TblRs and :AttRs opened,
'           ! where :TblRs is poiting the rec in tbl-att, if fnd just point to it, if not fnd, add one rec with Attn=@Attn
'           ! and   :AttRs is pointing to the :FileData of the fld-Attn of the tbl-Attn
Dim Q$: Q = FmtQQ("Select * from Att where Attn='?'", Attn)
If Not HasRecQ(D, Q) Then
    D.Execute FmtQQ("Insert into Att (Attn) values('?')", Attn) ' add rec to tbl-att with Attn=@Attn
End If
With Attd
    Set .Db = D
    Set .RecRs = Rs(D, Q)
    Dim F As DAO.Field2: Set F = .RecRs.Fields("Att")
    Dim V: Set V = .RecRs.Fields("Att").Value
    Set .FldRs = CvRs2(V) ' there is always a rec of Att=@Attn in .RecRs (Tbl-Attn)
End With
End Function

Function AttNy(D As Database) As String(): AttNy = SyzTF(D, "Att.Attn"): End Function
Function FileDataFd2zAttd(A As Attd) As DAO.Field2: Set FileDataFd2zAttd = A.FldRs!FileData: End Function

'**HasAtt
Function HasAtt(D As Database, Attn, Attf) As Boolean
Select Case True
Case Not HasAttn(D, Attn), Not HasAttf(D, Attn, Attf)
Case Else: HasAtt = True
End Select
End Function
Function NoAtt(D As Database, Attn, Attf) As Boolean:     NoAtt = Not HasAtt(D, Attn, Attf):                                 End Function
Function HasAttn(D As Database, Attn) As Boolean:       HasAttn = HasRecQ(D, FmtQQ("Select * from Att where Attn='?'", Attn)): End Function
Function HasAttf(D As Database, Attn, Attf) As Boolean: HasAttf = HasRecRsFeq(Attd(D, Attn).RecRs, "FileName", Attf):       End Function
Function HasRecTFeq(D As Database, T, F, Eqval) As Boolean

End Function

Private Sub AttFnAy__Tst(): D AttFnAyzNm(RelCstPgmDb, "AA"): End Sub
