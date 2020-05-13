Attribute VB_Name = "MxIdeCacMthFs"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthCacFs."
Const CNs$ = "Mth.Cac"

Private Function CacMthPth$(AssPth$)
CacMthPth = EnsPth(AssPth) & ".CacMth\"
End Function

Function CacMthDrsFtzM$(M As CodeModule)
CacMthDrsFtzM = CacMthDrsFt(AssPthzM(M), Mdn(M))
End Function

Function CacMthMD5FtzM$(M As CodeModule)
CacMthMD5FtzM = CacMthMD5Ft(AssPthzM(M), Mdn(M))
End Function

Function CacMthMD5FtM$()
CacMthMD5FtM = CacMthMD5FtzM(CMd)
End Function

Function CacMthDrsFtM$()
CacMthDrsFtM = CacMthDrsFtzM(CMd)
End Function

Sub BrwCacMthPthP()
BrwPth CacMthPthP
End Sub

Function CacMthPthP$()
CacMthPthP = CacMthPthzP(CPj)
End Function

Function CacMthPthzM$(M As CodeModule)
CacMthPthzM = CacMthPth(AssPthzM(M))
End Function

Function CacMthPthzP$(P As VBProject)
CacMthPthzP = CacMthPth(AssPthzP(P))
End Function

Private Function CacMthDrsFt$(AssPth$, Mdn$)
CacMthDrsFt = CacMthPth(AssPth) & Mdn & ".Drs.Txt"
End Function

Private Function CacMthMD5Ft$(AssPth$, Mdn$)
CacMthMD5Ft = CacMthPth(AssPth) & Mdn & ".Md5.Txt"
End Function

Sub EnsCacMthPthP()
EnsCacMthPthzP CPj
End Sub

Sub EnsCacMthPthzP(P As VBProject)
EnsAssPth Pjf(P)
EnsPth CacMthPthzP(P)
End Sub
