Attribute VB_Name = "MxDtaDaGpCnt"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaGpCnt."
Type GpCnt
    GpDy() As Variant ' Gp-Dy
    Cnt() As Long
End Type

Function GpCntzDy(Dy()) As GpCnt
'@D : :Drs-..{Gpcc}    ! it has columns-Gpcc
'Ret  :GpCnt  ! each ele-of-:GpCnt.Gp is a dr with fld as described by @Gpcc.  :GpCnt.Cnt is rec cnt of such gp
Dim OGpDy(), OCnt&()
    Dim Dr: For Each Dr In Itr(Dy)
        Dim Ix&: Ix = IxzDyDr(OGpDy, Dr)
        If Ix = -1 Then
            PushI OCnt, 1
            PushI OGpDy, Dr
        Else
            OCnt(Ix) = OCnt(Ix) + 1
        End If
    Next
GpCntzDy.GpDy = OGpDy
GpCntzDy.Cnt = OCnt
End Function

Function GpCntFny(D As Drs, GpFny$()) As GpCnt
'@D : :Drs-..{Gpcc}    ! it has columns-Gpcc
'Ret  :GpCnt  ! each ele-of-:GpCnt.Gp is a dr with fld as described by @Gpcc.  :GpCnt.Cnt is rec cnt of such gp
Dim OGpDy(), OCnt&()
    Dim A As Drs: A = DrszSelFny(D, GpFny)
    Dim I%: I = Si(A.Fny)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Ix&: Ix = IxzDyDr(OGpDy, Dr)
        If Ix = -1 Then
            PushI OCnt, 1
            PushI OGpDy, Dr
        Else
            OCnt(Ix) = OCnt(Ix) + 1
        End If
    Next
GpCntFny.GpDy = OGpDy
GpCntFny.Cnt = OCnt
End Function
Function GpCntAllCol(D As Drs) As GpCnt
GpCntAllCol = GpCntzDy(D.Dy)
End Function
Function GpCnt(D As Drs, Gpcc$) As GpCnt
GpCnt = GpCntFny(D, SyzSS(Gpcc))
End Function
