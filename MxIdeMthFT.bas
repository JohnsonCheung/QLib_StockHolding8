Attribute VB_Name = "MxIdeMthFT"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthFT."

Function MthBieyzN(Src$(), Mthn, Optional ShtMthTy$) As Bei()
Dim Ix: For Each Ix In Itr(MthixyzSNT(Src, Mthn, ShtMthTy))
   PushBei MthBieyzN, Bei(Ix, MthEix(Src, Ix))
Next
End Function

Function MthBeiy(Src$()) As Bei()
Dim Ix: For Each Ix In MthixItr(Src)
    PushBei MthBeiy, Bei(Ix, SrcEix(Src, Ix))
Next
End Function

Function MthBei(Src$(), Mthn, Optional ShtMthTy$) As Bei
Dim B&: B = Mthix(Src, Mthn, ShtMthTy)
MthBei = Bei(B, MthEix(Src, B))
End Function
