Attribute VB_Name = "MxDaoTblOpRen"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoTblOpRen."

Sub RenTbl(D As Database, T, ToNm$)
D.TableDefs(T).Name = ToNm
End Sub

Sub RenTblzFmPfx(D As Database, FmPfx$, ToPfx$)
Dim T As TableDef: For Each T In D.TableDefs
    If HasPfx(T.Name, FmPfx) Then
        T.Name = RplPfx(T.Name, FmPfx, ToPfx)
    End If
Next
End Sub

Sub RenTTzAddPfx(D As Database, TT$, Pfx$)
Dim T: For Each T In Ny(TT)
    RenTblzAddPfx D, CStr(T), Pfx
Next
End Sub
