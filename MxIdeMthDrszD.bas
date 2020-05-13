Attribute VB_Name = "MxIdeMthDrszD"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthDrszD."

Function DbCacMthDrs(D As Database) As Drs: DbCacMthDrs = Add6MthCol(DrszT(D, "Mth")): End Function
Function CDbCacMthDrs() As Drs: CDbCacMthDrs = DbCacMthDrs(CurrentDb): End Function
Function DbCacMthcDrs(D As Database, Pjn$, Mdny$()) As Drs
Dim O As Drs
Dim J%, N: For Each N In Itr(Mdny)
    Dim R As DAO.Recordset: Set R = Rs(D, FmtQQ("Select MdTy,Mdl from Md where Pjn='?' and Mdn='?'", Pjn, N))
    Dim Src$(): Src = SplitCrLf(CStr(R!Mdl))
    Dim T$: T = R!MdTy
    Stop
    Dim Dr(): 'Dr = MdnDr(Pjn, T, N)
    O = AddDrs(O, MthcDrszS(Src, Dr))
    J = J + 1
Next
DbCacMthcDrs = O
End Function

Function MdlzTbMdP$(Mdn): MdlzTbMdP = MdlzTbMd(CurrentDb, CPjn, Mdn): End Function

Function MdlzTbMd$(D As Database, Pjn$, Mdn)
Dim B$: B = FmtQQ("Pjn='?' and Mdn='?'", Pjn, Mdn)
MdlzTbMd = VzTF(D, "Md.Mdl", B)
End Function
