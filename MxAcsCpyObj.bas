Attribute VB_Name = "MxAcsCpyObj"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxAcsCpyObj."

Private Sub CpyAllFbObj__Tst()
Dim Fb$: Fb = StkHld8TmpFba
DltFfnIf Fb
CrtFb Fb
CpyAllFbObj StkHld8Fba, Fb
CpyAllFbObj PjfP, Fb, "QLib_"
Debug.Print Fb
End Sub

Sub CpyAllFbObj(FmFb$, ToFb$, Optional NewObjPfx$)
Dim A As Access.Application: Set A = AcszFb(FmFb, IsExl:=True)
ClsAllFrm A

CpyAllAcsRf A, ToFb
CpyAllAcsTbl A, ToFb, NewObjPfx
CpyAllAcsFrm A, ToFb, NewObjPfx
CpyAllAcsQry A, ToFb, NewObjPfx
CpyAllAcsMd A, ToFb, NewObjPfx
CpyAllAcsRpt A, ToFb, NewObjPfx
QuitAcs A
End Sub

Sub CpyAllAcsFrm(A As Access.Application, ToFb$, Optional NewObjPfx$)
ClsAllFrm A
Dim F: For Each F In Itr(FrmNy(A.CurrentDb))
    A.DoCmd.CopyObject ToFb, NewObjPfx & F, acForm, F
Next
End Sub

Sub CpyAllAcsRpt(A As Access.Application, ToFb$, Optional NewObjPfx$)
Dim R: For Each R In Itr(RptNy(A.CurrentDb))
    A.DoCmd.CopyObject ToFb, NewObjPfx & R, acReport, R
Next
End Sub

Sub CpyAllAcsMd(A As Access.Application, ToFb$, Optional NewMdPfx$)
Dim M: For Each M In Mdny(MainPj(A))
    A.DoCmd.CopyObject ToFb, NewMdPfx & M, acModule, M
Next
End Sub

Sub CpyAllAcsTbl(A As Access.Application, ToFb$, Optional NewObjPfx$)
Dim LTbl$(), Nrm$(), LCnn$(), LSrc$()
    Dim FmD As Database: Set FmD = A.CurrentDb
    Dim Td As DAO.TableDef
    Dim T: For Each T In Itr(Tny(FmD))
        Set Td = FmD.TableDefs(T)
        If IsTdLnk(Td) Then
            PushI LTbl, Td.Name
            PushI LCnn, Td.Connect
            PushI LSrc, Td.SourceTableName
        Else
            PushI Nrm, T
        End If
    Next
Dim ToDb As Database: Set ToDb = Db(ToFb)
    Dim J%: For J = 0 To UB(LTbl)
        LnkTbl ToDb, LTbl(J), LSrc(J), LCnn(J)
    Next
    ToDb.Close
For Each T In Itr(Nrm)
    A.DoCmd.CopyObject ToFb, NewObjPfx & T, acTable, T
Next
End Sub

Sub CpyAllAcsQry(A As Access.Application, ToFb$, Optional NewQryPfx$)
Dim I As QueryDef: For Each I In A.CurrentDb.QueryDefs
    A.DoCmd.CopyObject ToFb, NewQryPfx & I.Name, acQuery, I.Name
Next
End Sub

Sub CpyAllAcsRf(A As Access.Application, ToFb$)
Dim ToAcs As Access.Application: Set ToAcs = AcszFb(ToFb)
Dim ToPj As VBProject: Set ToPj = MainPj(ToAcs)
Dim FmPj As VBProject: Set FmPj = MainPj(A)
Dim R As VBIde.Reference: For Each R In FmPj.References
    Debug.Print R.Name
    If NoRf(ToPj, R.Name) Then
        ToPj.References.AddFromFile R.FullPath
    End If
Next
QuitAcs ToAcs
End Sub
