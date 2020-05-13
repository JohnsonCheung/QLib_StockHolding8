Attribute VB_Name = "MxDtaDaDyAli"
Option Explicit
Option Compare Text
Const CNs$ = "Dy.Op.Ali"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaDyAli."

Function AliDr$(Dr, Wdty%())
'Ret : a Ln by joing [ | ] and quoting [| * |] after aligng @Dr with @Wdty. @@
AliDr = QuoJnzAsTLn(AliDrzW(Dr, Wdty))
End Function

Function AliDrzW(Dr, Wdty%()) As Variant()
Dim O()
Dim J%: For J = 0 To Min(UB(Dr), UB(Wdty))
    PushI O, AliL(Dr(J), Wdty(J))
Next
AliDrzW = O
End Function

Sub AliDyzCol(ODy(), C)
'Fm ODy : the col @C will be aligned
'Fm C    : the column ix
'Ret     : column-@C of @ODy will be aligned
Dim Col(): Col = ColzDy(ODy, C)
Dim ACol$(): ACol = AmAli(Col)
Dim J&: For J = 0 To UB(ODy)
    ODy(J)(C) = ACol(J)
Next
End Sub


Function AliDyzCix(Dy(), Cix&()) As Variant()
Dim O(): O = Dy
Dim C: For Each C In Cix
    AliDyzCol O, C
Next
AliDyzCix = O
End Function

Function AliDy(Dy(), Optional FstNTerm%, Optional RstWdt% = 120) As Variant()
Dim W%(): W = AddAyEle(WdtyzDy(Dy, FstNTerm), RstWdt)
AliDy = AliDyzW(Dy, W)
End Function

Function AliDyAsLy(Dy()) As String()
AliDyAsLy = JnDy(AliDy(Dy))
End Function

Function AliDyzW(Dy(), FstNTermWdty%()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI AliDyzW, AliDrzW(Dr, FstNTermWdty)
Next
End Function
