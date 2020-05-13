Attribute VB_Name = "MxIdeMthLisDrs"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Mth.Lis"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthLisDrs."

Function MthLisDrs() As Drs
MthLisDrs = MthLisDrszP(CPj)
End Function

Function MthLisDrszP(P As VBProject) As Drs
MthLisDrszP = MthLisDrszMth(MthDrszP(P))
End Function

Function MthLisDrszMth(MthDrs As Drs) As Drs
MthLisDrszMth = SelDrs(Add5MthCol(MthDrs), MthFFLis)
End Function

Function MthLisWh(MthLisDrs As Drs, W As WhMth) As Drs
'- Pfx-Pn = Patn
Dim PnMdy$:             PnMdy = PatnzWhSS(W.ShtMdySS, ShtMdyAy)
Dim PnTy$:               PnTy = PatnzWhSS(W.ShtMthTySS, ShtMthTyAy)
'- Pfx-I = Inp-Do-Fm-MthLisDrs
Dim IMdy   As Drs:     IMdy = DwPatn(MthLisDrs, "Mdy", PnMdy)
Dim ITy    As Drs:      ITy = DwPatn(IMdy, "Ty", PnTy)
Dim ITyChr As Drs:   ITyChr = DwEqStr(ITy, "TyChr", W.TyChr)
Dim IPatn  As Drs:    IPatn = DwPatn(ITyChr, "Mthn", W.Patn)
Dim IHasAp As Drs:   IHasAp = DwAnyAp(IPatn, W.AnyAp)
Dim INPm   As Drs:     INPm = DwNPm(IHasAp, W.NPm)
Dim IMdn   As Drs:     IMdn = DwPatn(INPm, "Mdn", W.MdnPatn)
Dim IRetAs As Drs:   IRetAs = DwPatn(IMdn, "RetAs", W.RetAsPatn)
Dim IRetAy As Drs:   IRetAy = DwRetAy(IRetAs, W.RetAy)
                   MthLisWh = DwPatn(IRetAy, "ShtPm", W.ShtPmPatn)
End Function

Private Function DwRetAy(WiRetAs As Drs, RetAy As eTri) As Drs
If RetAy = eTriOpn Then DwRetAy = WiRetAs: Exit Function
Dim RetAy1 As Boolean: RetAy1 = BoolzTri(RetAy)
Dim IRetAs%: IRetAs = IxzAy(WiRetAs.Fny, "RetAs")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiRetAs.Dy)
        Dim RetAs$: RetAs = Dr(IRetAs)
        If HasSfx(RetAs, "()") = RetAy1 Then PushI ODy, Dr
    Next
DwRetAy = Drs(WiRetAs.Fny, ODy)
End Function

Private Function HasAp(MthPm) As Boolean
Dim A$(): A = SplitCommaSpc(MthPm): If Si(A) = 0 Then Exit Function
HasAp = HasPfx(LasEle(A), "ParamArray ")
End Function

Private Function DwAnyAp(WiMthPm As Drs, HasAp0 As eTri) As Drs
If HasAp0 = eTriOpn Then DwAnyAp = WiMthPm: Exit Function
Dim HasAp1 As Boolean: HasAp1 = BoolzTri(HasAp0)
Dim IMthPm%: IMthPm = IxzAy(WiMthPm.Fny, "MthPm")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiMthPm.Dy)
        Dim MthPm$: MthPm = Dr(IMthPm)
        If HasAp1 = HasAp(MthPm) Then PushI ODy, Dr
    Next
DwAnyAp = Drs(WiMthPm.Fny, ODy)
End Function

Private Function DwNPm(D As Drs, NPm%) As Drs
If NPm < 0 Then DwNPm = D: Exit Function
Dim Ix%: Ix = IxzAy(D.Fny, "MthPm")
Dim ODy(), Dr, Pm$: For Each Dr In Itr(D.Dy)
    Pm = Dr(Ix)
    If Si(SplitComma(Pm)) = NPm Then PushI ODy, Dr
Next
DwNPm = Drs(D.Fny, ODy)
End Function
