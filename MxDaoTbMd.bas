Attribute VB_Name = "MxDaoTbMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDaoTbMd."

Sub RfhTbMd(D As Database, RfhId&, P As VBProject)
'Assume: @D has TbMd-
'@D-Md: Pjn Mdn MdTy Lib Mdl IsNotUse CrtId UpdId Si NLin
Dim Pjn$: Pjn = P.Name
Dim OldN$(): OldN = StrColzTF(D, "Md", "Mdn")  ' #Old-MdNy#   ! $$Md-Mdn-Ay
Dim CurN$(): CurN = Mdny(P)                 ' #Cur-MdNy#   ! @P->Modules-Ny
Dim ExiN$(): ExiN = IntersectAy(OldN, CurN)    ' #Exist-MdNy# ! MdNy In both $$Md & @P
Dim ExiL$(): ExiL = MdlAy(P, ExiN)            ' #Exist-Mdl#  ! Mdl  in both $$Md & @P

Dim NewN$(): NewN = MinusAy(CurN, OldN)        ' #New-MdNy#   ! In @P not in $$Md
Dim Newl$(): Newl = MdlAy(P, NewN)         ' #New-Mdl#    ! In @P not in $$Md
Dim NewT$(): NewT = ShtCmpTyAy(P, NewN)
Dim DltN$(): DltN = MinusAy(OldN, CurN)        ' #NotUsed-MdNy ! in $$Md not in in @P
    
InsNew RfhId, Pjn, NewN, Newl, NewT
RmkNotUse RfhId, Pjn, DltN
RfhDif RfhId, Pjn, ExiN
End Sub

Private Sub RfhDif(RfhId&, Pjn$, ExiN$())
Dim Newl$, Oldl$
Dim NewT$, OldT$
Dim R As DAO.Recordset: Set R = CurrentDb.TableDefs("Md").OpenRecordset
With R
    .Index = "PrimaryKey"

    Dim N: For Each N In Itr(ExiN)
        Dim M As CodeModule: Set M = Md(N)
        Newl = Mdl(M)
        NewT = ShtCmpTy(CmpTyzM(M))
        
        .Seek "=", Pjn, N
        If .NoMatch Then Stop
        Oldl = !Mdl
        OldT = Nz(!MdTy, "")
        If _
            Newl <> Oldl Or _
            NewT <> OldT Then
            
            .Edit
            !Mdl = Oldl
            !MdTy = NewT
            !UpdId = RfhId
            !IsNotUse = False
            !Si = Len(Oldl)
            !NLin = LnCnt(Oldl)
            .Update
    
        End If
    Next
End With
End Sub


Private Sub InsNew(RfhId&, Pjn$, NewN$(), Newl$(), NewT$())
'@NewN : ! New Md name to be inserted
'@NewL : ! New Md Lines to be inserted
'@NewT : ! New Md Sht Ty to be inserted
'Do : ! Ins new Md into TbMd as indicated in 5 given parameters.
'     ! where TbMd: Pjn Mdn MdTy Lib Mdl IsNotUse CrtId UpdId Si NLin

Dim R As DAO.Recordset: Set R = CurrentDb.TableDefs("Md").OpenRecordset
Dim J%: For J = 0 To UB(NewN)
    R.AddNew
    R!Pjn = Pjn
    R!Lib = CLibv(DclzSrc(SplitCrLf(Newl(J))))
    R!Mdn = NewN(J)
    R!MdTy = NewT(J)
    R!Mdl = Newl(J)
    R!CrtId = RfhId
    R!UpdId = RfhId
    R!Si = Len(R!Mdl)
    R!NLin = LnCnt(R!Mdl)
    R.Update
Next
R.Close
End Sub

Private Sub RmkNotUse(RfhId&, Pjn$, DltN$())
With CurrentDb.TableDefs("Md").OpenRecordset
    .Index = "PrimaryKey"
    Dim J%: For J = 0 To UB(DltN)
        .Seek "=", Pjn, DltN(J)
        If .NoMatch Then Stop
        .Edit
        !UpdId = RfhId
        !IsNotUse = True
        .Update
    Next
    .Close
End With
End Sub

Sub RfhCTbMd()
RfhTbMd CurrentDb, NwRfhId(CurrentDb), CPj
End Sub
