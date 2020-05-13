Attribute VB_Name = "MxVbStrPfxSfx"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrPfxSfx."

Function AddStrAp$(ParamArray StrAp())
Dim Av(): Av = StrAp
AddStrAp = Jn(Av)
End Function

Function AddPfx(S, Pfx): AddPfx = Pfx & S: End Function
Function AddPfxS(S, Pfx, Sfx): AddPfxS = Pfx & S & Sfx: End Function

Function IsAllDig(S) As Boolean
Dim J%: For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsAllDig = True
End Function

Function IsNDig(S, N%) As Boolean
Dim L%: L = Len(S)
Select Case True
Case L <> N
Case Not IsAllDig(S)
Case Else: IsNDig = True
End Select
End Function


Function AddSfx(S, Sfx): AddSfx = S & Sfx: End Function


Function HasPfxzSomSyEle(Sy$(), Pfx) As Boolean
Dim I: For Each I In Itr(Sy)
   If Not HasPfx(I, Pfx) Then Exit Function
Next
HasPfxzSomSyEle = True
End Function

Function SfxChr$(S, SfxChrLis$, Optional C As VbCompareMethod = vbBinaryCompare)
If HasSfxChr(S, SfxChrLis, C) Then SfxChr = LasChr(S)
End Function

Function Sfx$(S, Suffix$, Optional C As VbCompareMethod = vbBinaryCompare)
If HasSfx(S, Suffix, C) Then Sfx = Suffix
End Function

Function HasSfxChr(S, SfxChrLis$, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
Dim J%
For J = 1 To Len(SfxChrLis)
    If HasSfx(S, Mid(SfxChrLis, J, 1), C) Then HasSfxChr = True: Exit Function
Next
End Function
Function HasPfxOfAllEle(Ay, Pfx, Optional C As eCas) As Boolean
If Si(Ay) = 0 Then Exit Function
Dim V: For Each V In Itr(Ay)
    If Not HasPfx(V, Pfx, C) Then Exit Function
Next
HasPfxOfAllEle = True
End Function

Function NoPfx(S, Pfx, Optional C As eCas) As Boolean: NoPfx = Not HasPfx(S, Pfx, C): End Function
Function HasPfx(S, Pfx, Optional C As eCas) As Boolean: HasPfx = StrComp(Left(S, Len(Pfx)), Pfx, CprMth(C)) = 0: End Function
Function HasPfxSfx(S, Pfx, Sfx, Optional C As VbCompareMethod = vbTextCompare) As Boolean
If Not HasPfx(S, Pfx, C) Then Exit Function
If Not HasSfx(S, Sfx, C) Then Exit Function
HasPfxSfx = True
End Function

Function HasPfxss(S, Pfxss$, Optional C As VbCompareMethod = vbTextCompare) As Boolean
Dim Pfxy$(): Pfxy = SyzSS(Pfxss)
HasPfxss = HasPfxy(S, Pfxy, C)
End Function
Function HasPfxy(S, Pfxy$(), Optional C As VbCompareMethod = vbTextCompare) As Boolean
Dim Pfx: For Each Pfx In Itr(Pfxy)
    If HasPfx(S, Pfx, C) Then HasPfxy = True: Exit Function
Next
End Function

Function HasPfxzAy(Ay, Pfx, Optional C As eCas) As Boolean
Dim I: For Each I In Itr(Ay)
    If HasPfx(I, Pfx, C) Then HasPfxzAy = True: Exit Function
Next
End Function

Function HasSfx(S, Sfx, Optional C As eCas) As Boolean: HasSfx = IsEqStr(Right(S, Len(Sfx)), Sfx, C): End Function
Function HasSfxApIgnCas(S, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
HasSfxApIgnCas = HasSfxAv(S, Av, vbTextCompare)
End Function

Function HasSfxAv(S, SfxAv(), C As VbCompareMethod) As Boolean
Dim Sfx
For Each Sfx In SfxAv
    If HasSfx(S, Sfx, C) Then HasSfxAv = True: Exit Function
Next
End Function

Function PfxzAy$(S, Pfxy) 'Pfx of Ay :: Pfx of @S has a pfx in @Pfxy
Dim P: For Each P In Pfxy
    If HasPfx(S, P) Then PfxzAy = P: Exit Function
Next
End Function

Function PfxzAySpc$(S, Pfxy$()) ' Pfx of Ay+S :: ret one of @Pfxy if @S has such pfx+space or blnk
Dim P: For Each P In Pfxy
    If HasPfx(S, P & " ") Then PfxzAySpc = P: Exit Function
Next
End Function

Function SfxzAySpc$(S, Sfxy$()) ' Sfx of Ay+S :: return one of @Sfxy if @S has such sfx+spc or blnk
Dim Sfx: For Each Sfx In Sfxy
    If HasSfxSpc(S, Sfx) Then SfxzAySpc = Sfx: Exit Function
Next
End Function

Function HasSfxSpc(S, Sfx, Optional C As eCas) As Boolean
HasSfxSpc = HasSfx(S, Sfx + " ", C)
End Function

Function PfxzSpc$(S, Pfx$) ' Pfx of spc :: return @Pfx if @S has @Pfx+spc or blnk
If HasPfx(S, Pfx & " ") Then PfxzSpc = Pfx
End Function

Function Pfx$(S, P$) ' Pfx :: return @P if @S has Pfx-@P or blnk
If HasPfx(S, P) Then Pfx = P
End Function

Function PfxzAp(S, ParamArray PfxAp()) ' Pfx of Ap :: ret one of @PfxAp if @S has such pfx or blnk
Dim PfxAv(): PfxAv = PfxAp
PfxzAp = PfxzAy(S, PfxAv)
End Function
