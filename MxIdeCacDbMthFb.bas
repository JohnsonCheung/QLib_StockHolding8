Attribute VB_Name = "MxIdeCacDbMthFb"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthFb."
Function MthFbP$()
MthFbP = MthFbzP(CPj)
End Function

Function MthFb$()
MthFb = MthFbP
End Function

Function MthFbzP$(P As VBProject)
MthFbzP = MthPthzP(P) & Fn(Pjf(P)) & ".MthDb.accdb"
End Function

Function MthPthzP$(P As VBProject)
Dim F$: F = Pjf(P)
MthPthzP = AddFdrEns(Pth(F), ".MthDb")
End Function

Function EnsMthFb(MthFb$) As Database
EnsFb MthFb
Dim D As Database
Set EnsMthFb = Db(MthFb)
'EnsSchm D, LnoChm
End Function

Function MthDbzP(P As VBProject) As Database
Dim Fb$: Fb = MthFbzP(P)
EnsMthFb Fb
Set MthDbzP = Db(Fb)
End Function
 
Function MthDbP() As Database
Set MthDbP = MthDbzP(CPj)
End Function

Sub BrwMthFb()
BrwFb MthFb
End Sub

Property Get LnoChm() As String()
Erase XX
X "Fld"
X " Nm  Md Pj"
X " T50 MchStr"
X " T10 MthPfx"
X " Txt Pjf Prm Ret LinRmk"
X " T3  Ty Mdy"
X " T4  MdTy"
X " Lng Lno"
X " Mem Lines Mrmk"
X "Tbl"
X " Pj  *Id Pjf | Pjn PjDte"
X " Md  *Id PjId Mdn | MdTy"
X " Mth *Id MdId Mthn ShtTy | ShtMdy Prm Ret LinRmk Mrmk Lines Lno"
LnoChm = XX
Erase XX
End Property
