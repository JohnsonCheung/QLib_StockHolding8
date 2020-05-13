Attribute VB_Name = "MxDtaDaGDrs"
Option Explicit
Option Compare Text
Const CNs$ = "Dta.Op"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaGDrs."

Function GDrs(D As Drs, Keycc$, Gpcc$) As Drs
'@D     : :Drs ..Keycc ..Gpcc: ! with fields as described in @Keycc @Gpcc
'@Keycc : :CC                  ! #Key-Column-Names# These columns in @D will be returned as first #NKey columns of returned @@Drs
'@Gpcc  : :CC                  ! #Gp-Column-Names# These columns in @D will be grouped as last-field of returned $$Drs @@
':GDrs: :Drs ! #Grouped-Drs# it is a Drs with #NKey + 1 columns
'            ! whose Fny is with first #NKey is from @Keycc
'             ! .    last column name is in format of "Gp-<Gpc1>-<Gpc2>-..-<GpcN>"
'                              !              which Gpc<i> is coming from @Gpcc

Dim KeyDrs As Drs: KeyDrs = SelDrs(D, Keycc)
Dim KeyGRxy():    KeyGRxy = GRxy(KeyDrs.Dy)
Dim DistKeyDy() ': Dim KeyGRxy(), KeyDrs As Drs
    Dim W1Dy(): W1Dy = KeyDrs.Dy
    Dim W1IRxy: For Each W1IRxy In Itr(KeyGRxy)
        Dim W1Rxy&(): W1Rxy = W1IRxy
        PushI DistKeyDy, W1Dy(W1Rxy(0))   '<---
    Next
Dim GpDy(): GpDy = SelDrs(D, Gpcc).Dy
Dim ODy(): 'Dim KEyGRxy(), GpDy(), DistKeyDy()
    Dim W2IxDistKey%: W2IxDistKey = 0
    Dim W2IRxy: For Each W2IRxy In Itr(KeyGRxy)
        Dim W2IGpDy()
            Erase W2IGpDy
            Dim W2IRix: For Each W2IRix In W2IRxy
                PushI W2IGpDy, GpDy(W2IRix)
            Next
        Dim W2Dr(): W2Dr = DistKeyDy(W2IxDistKey)
                           PushI W2Dr, W2IGpDy
                           PushI ODy, W2Dr '<--
             W2IxDistKey = W2IxDistKey + 1
    Next
Dim OFny$(): 'Dim KeyDrs As Drs
    OFny = KeyDrs.Fny
           PushI OFny, "Gp-" & Replace(Gpcc, " ", "-")
           
GDrs = Drs(OFny, ODy)
'BrwDrs GDrs: Stop
End Function
