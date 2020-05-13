Attribute VB_Name = "MxIdeSrcDclCnstv"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Dcl.3Cnst"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclCnstv."
':CNsv: :S  ! #Cnst-CNs-Value#  PrimFm-Dcl. the string bet-DblQ of CnstLin-CNs of a Md
':CModv: :S ! #Cnst-CMod-Value# PrimFm-Dcl. the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CMod of a Md
':CLibv: :S ! #Cnst-CLib-Value# PrimFm-Dcl. the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CLib of a Md

Function CNsLin$(Ns$)
':CLibLin: :PrvCnstLin ! Is a `Const CLib$ = "${Clibv}."`
If Ns = "" Then Exit Function
CNsLin = FmtQQ("Const CNs$ = ""?""", Ns)
End Function

Function CLibLin$(CLibv$)
':CLibLin: :PrvCnstLin ! Is a `Const CLib$ = "${Clibv}."`
CLibLin = FmtQQ("Const CLib$ = ""?.""", CLibv)
End Function

Function CModLin$(M As CodeModule)
':CModLin: :CnstLin ! Is a Const CMod$ = CLib & "xxxx."
CModLin = FmtQQ("Const CMod$ = CLib & ""?.""", Mdn(M))
End Function

Function CNsvzM$(M As CodeModule)
CNsvzM = CNsv(Dcl(M))
End Function

Private Sub NsAyP__Tst()
BrwAy NsAyP
End Sub

Function NsAyP() As String()
NsAyP = NsAyzP(CPj)
End Function

Function NsAyzP(P As VBProject) As String()
Dim O$()
Dim C As VBComponent: For Each C In P.VBComponents
    PushNBNDup O, CNsv(Dcl(C.CodeModule))
Next
NsAyzP = SrtAy(O)
End Function

Function CNsvM$()
CNsvM = CNsvzM(CMd)
End Function

Function CNsv$(Dcl$())
CNsv = StrCnstv(Dcl, "CNs")
End Function

Function CModv$(Dcl$())
CModv = RmvSfxDot(StrCnstv(Dcl, "CMod"))
End Function

Function HasCLibv(Dcl$(), Libn$) As Boolean
HasCLibv = CLibv(Dcl) = Libn
End Function

Function CLibv$(Dcl$())
CLibv = RmvLasChr(StrCnstv(Dcl, "CLib"))
End Function

Function CLibvzM$(M As CodeModule)
CLibvzM = RmvLasChr(StrCnstvzM(M, "CLib"))
End Function

Function CLibvM$()
CLibvM = CLibvzM(CMd)
End Function

Function CLibvAyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushNDup CLibvAyzP, CLibv(Dcl(C.CodeModule))
Next
End Function

Function CLibvAyP() As String()
CLibvAyP = CLibvAyzP(CPj)
End Function
Function LibNyP() As String()
LibNyP = LibNyzP(CPj)
End Function

Function LibNyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushNBNDup LibNyzP, CLibvzM(C.CodeModule)
Next
End Function
