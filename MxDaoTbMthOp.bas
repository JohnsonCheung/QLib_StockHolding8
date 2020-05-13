Attribute VB_Name = "MxDaoTbMthOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDaoTbMthOp."

Sub RfhCTbMth()
RfhTbMth CurrentDb
End Sub

Private Sub RfhTbMthzId__Tst()
RfhTbMthzId CurrentDb, LasRfhId(CurrentDb), CPj
End Sub

Sub RfhTbMthzId(D As Database, RfhId&, Pj As VBProject)
'Do : delete reocrds to @D.Mth for those record @RfhId by
'     insert records to @D.Mth from $$Md->Mdl
Dim Ny$(), Ty$(): AsgStrColPair StrColPair(D, "Md", "Mdn MdTy", "UpdId=" & RfhId), Ny, Ty
Dim Pjn$: Pjn = CPjn
Dim N: For Each N In Itr(Ny)
    D.Execute FmtQQ("Delete * from Mth where Mthn='?' and Pjn='?'", N, Pjn)
Next
'InsTblzDrs D, "Mth", MthDrszD(D, Pjn, ny)
End Sub

Sub RfhTbMth(D As Database)
Dim Pj As VBProject: Set Pj = CPj
Dim RfhId&: RfhId = NwRfhId(D)
RfhTbMd D, RfhId, Pj
RfhTbMthzId D, RfhId, Pj

'Upd $$Lib from $$Md
RunCQ "Delete * from Lib"
RunCQ "Insert Into Lib Select Distinct Lib from Md"

'Upd $$Pj from $$Md
RunCQ "Delete * from Pj"
RunCQ "Insert Into Lib Select Distinct Pj from Md"
End Sub

Function LasRfhId&(D As Database)
LasRfhId = LasId(D, "RfhHis")
End Function

Function NwRfhId&(D As Database)
With D.TableDefs("RfhHis").OpenRecordset
    .AddNew
    NwRfhId = !RfhId
    .Update
    .Close
End With
End Function

Private Sub InsTbMth(D As Database, Pjn$, Mdny$(), ShtMdTy$())
Dim Ns$
Dim N, J&: For Each N In Itr(Mdny)
    Dim Src$(): Src = SplitCrLf(MdlzTbMd(D, Pjn, N))
    Dim Dr(): Dr = MdnDr(Pjn, Ns, ShtMdTy(J), CStr(N), Si(Src))
    Dim Drs As Drs: Drs = MthcDrszS(Src, Dr)
    Stop 'InsTblzDrs D, "Mth", MthDrs(D, Pjn, Mdny)
    J = J + 1
Next
End Sub
