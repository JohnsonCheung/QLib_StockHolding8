Attribute VB_Name = "MxIdeSrcGenLibp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcGenLibp."

Function LibSrcPthzP$(P As VBProject, Libv$)
':LibSrcPth: :Pth #Library-Src-Pth# ! a :SrcPth under :Libp.
'                                 ! under this :Libp, there are 2-or-more SrcPth with Fdr-name = `{Libv}.{Ext}.src`, where {Ext} is Ext-of-Pjf-of-@P.
'                                 ! Note: The folder of the SrcPth is in format of `{PjFn}.src`
'                                 ! Example: Given Pjf-@P            "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb"
'                                 !          Given @Libv             "QVb"
'                                 !          Then LibpzP(@P) will be "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb.lib\QVb.accdb.src"
Dim LibPth$: LibPth = LibpzP(P)
Dim LibExt$: LibExt = Ext(RmvExt(RmvPthSfx(LibPth)))
Dim LibFdr$: LibFdr = Libv & LibExt & ".src"
LibSrcPthzP = EnsAllFdr(LibPth & LibFdr)
End Function

Sub EnsLibSrcPth(P As VBProject, Libv$)
EnsAllFdr LibSrcPthzP(P, Libv)
End Sub

Function LibSrcPthzDistPj$(DistPj As VBProject)
Dim P$: P = Pjp(DistPj)
LibSrcPthzDistPj = AddFdrAp(UpPth(P, 1), ".Src", Fdr(P))
End Function

Function LibSrcPthP$(Libv$)
LibSrcPthP = LibSrcPthzP(CPj, Libv)
End Function

Function LibpP$()
LibpP = LibpzP(CPj)
End Function

Function LibpzP$(P As VBProject)
':Libp: :Pth #Library-Pth# ! a pj can generate 1 pj or 2-or-more-pj.  When gen 2-or-more-pj, there is a :Libp in the same fdr as the Pjf of @P.
'                          ! under this :Libp, there are 2-or-more SrcPth with Fdr-name = `{Libv}.{Ext}.src`, where {Ext} is Ext-of-Pjf-of-@P.
'                          ! Note: The folder of the SrcPth is in format of `{PjFn}.src`
'                          ! Example: Given Pjf-@P            "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb"
'                          !          Then LibpzP(@P) will be "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb.lib\"
LibpzP = EnsPth(Pjf(P) & ".lib")
End Function

Sub BrwLibpP()
BrwPth LibpP
End Sub

Function IsLibSrcPth(Pth) As Boolean
Dim F$: F = Fdr(Pth)
If Not HasExtSS(F, ".xlam .accdb") Then Exit Function
IsLibSrcPth = Fdr(ParPth(Pth)) = ".Src"
End Function
