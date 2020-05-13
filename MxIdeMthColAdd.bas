Attribute VB_Name = "MxIdeMthColAdd"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Drs.AddCol"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthColAdd."

Function Add5MthCol(WiMthln As Drs) As Drs
Dim ITyChr   As Drs:     ITyChr = AddMthColTyChr(WiMthln)
Dim IPm      As Drs:        IPm = AddMthColMthPm(ITyChr)
Dim IShtPm   As Drs:     IShtPm = AddMthColShtPm(IPm)
Dim IRetAs   As Drs:     IRetAs = AddMthColRetTyn(IShtPm)
Add5MthCol = AddMthColIsRetObj(IRetAs)
End Function

Function AddMthColTyChr(WiMthln As Drs) As Drs
'Ret         : Add col-HasPm
Dim I%: I = IxzAy(WiMthln.Fny, "Mthln")
Dim Dr, Dy(): For Each Dr In Itr(WiMthln.Dy)
    Dim Mthln$: Mthln = Dr(I)
    Dim TyChr$: TyChr = MthChr(Mthln)
    PushI Dr, TyChr
    PushI Dy, Dr
Next
AddMthColTyChr = AddColzFFDy(WiMthln, "TyChr", Dy)
End Function

Function AddMthColShtPm(WiMthPm As Drs) As Drs
'Ret         : Add col-ShtPm
Dim I%: I = IxzAy(WiMthPm.Fny, "MthPm")
Dim Dr, Dy(): For Each Dr In Itr(WiMthPm.Dy)
    Dim MthPm$: MthPm = Dr(I)
    Dim ShtPm1$: ShtPm1 = ShtMthPm(MthPm)
    PushI Dr, ShtPm1
    PushI Dy, Dr
Next
AddMthColShtPm = AddColzFFDy(WiMthPm, "ShtPm", Dy)
End Function
Function AddColzBetBkt(D As Drs, ColnAs$, Optional IsDrp As Boolean) As Drs
Dim BetColn$, NewC$: AsgBrk1 ColnAs, ":", BetColn, NewC
If NewC = "" Then NewC = BetColn & "InsideBkt"
Dim Ix%: Ix = IxzAy(D.Fny, BetColn)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    PushI Dr, BetBkt(Dr(Ix))
    PushI Dy, Dr
Next
Dim O As Drs: O = AddColzFFDy(D, NewC, Dy)
If IsDrp Then O = DrpColzDrsCC(O, BetColn)
AddColzBetBkt = O
End Function

Function Add6MthCol(MthDrs As Drs) As Drs

End Function
Function AddMthColMthPm(WiMthln As Drs, Optional IsDrp As Boolean) As Drs
AddMthColMthPm = AddColzBetBkt(WiMthln, "Mthln:MthPm", IsDrp)
End Function

Function AddMthColIsRetObj(WiRetAs As Drs) As Drs
'@WiRetAs :Drs..RetAs..
'Ret       :Drs..IsRetObj @@
Dim IxRetAs%: IxRetAs = IxzAy(WiRetAs.Fny, "RetAs")
Dim Dr, Dy(): For Each Dr In Itr(WiRetAs.Dy)
    Dim RetAs$: RetAs = Dr(IxRetAs)
    Dim R As Boolean: R = IsObjTyn(RetAs)
    PushI Dr, R
    PushI Dy, Dr
Next
AddMthColIsRetObj = AddColzFFDy(WiRetAs, "IsRetObj", Dy)
End Function

Function AddMthColRetTyn(WiMthln As Drs) As Drs
Dim I%: I = IxzAy(WiMthln.Fny, "Mthln")
Dim Dr, Dy(): For Each Dr In Itr(WiMthln.Dy)
    Dim Mthln$: Mthln = Dr(I)
    Dim Ret$: Ret = RetTyn(Mthln)
    PushI Dr, Ret
    PushI Dy, Dr
Next
AddMthColRetTyn = AddColzFFDy(WiMthln, "RetAs", Dy)
End Function

Function RetTyn$(Mthln)
RetTyn = ShfNmAftAs(AftBkt(Mthln))
End Function

