Attribute VB_Name = "MxVbStrCml"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrCml."
':Cml: :Nm #Camel#         ! :: [ Cmlf | Cmln | Cmll ] [Cmln].. ::
':Cmlf: :Nm #Camel<fat>#    ! FstNChar is UCase, rest is :CmlRestChr
':Cmln: :Nm #Camel<normal># ! FstChr is UCase, rest is :CmlRestChr
':Cmll: :Nm #Camel<lower-case># ! FstChr is Lcase, rest is :CmlRestChr
':CmlRestChr: :Chr #Camel-Rest-Chr# ! Lower-Case or _ or digit

Private Sub ShfCml__Tst()
Dim L$, EptL$
Ept = "A"
L = "AABcDD"
EptL = "ABcDD"
GoSub Tst
Exit Sub
Tst:
    Act = ShfCml(L)
    If Act <> Ept Then Stop
    If EptL <> L Then Stop
    Return
End Sub

Function CmlItr(Nm)
CmlItr = Itr(CmlAy(Nm))
End Function

Private Sub CmlAy__Tst()
Dim Nm$
GoSub YY
Exit Sub
Z1:
    Dim Ny$(): Ny = MthnyV
    Dim N: For Each N In Ny
        If N <> Jn(CmlAy(CStr(N))) Then Stop
    Next
    Return
YY:
    Nm = "A_IxAy"
    Ept = Sy("A_", "Ix", "Ay")
    GoTo Tst
Tst:
    Act = CmlAy(Nm)
    C
    Return
End Sub

Function CapCmlAy(Nm) As String()
Dim M$, N$, J%
N = Nm
Again:
    LoopTooMuch CSub, J
    M = ShfCapCml(N)
    If M = "" Then Exit Function
    PushI CapCmlAy, M
    GoTo Again
End Function

Function CmlAy(Nm) As String()
Dim S$: S = Nm
While S <> ""
    PushI CmlAy, ShfCml(S)
Wend
End Function

Function ShfCml$(ONm)
Const CSub$ = CMod & "ShfCml"
If ONm = "" Then Exit Function
Dim O$
Again:
    Dim M$: M = ShfCapCml(ONm)
    If M = "" Then GoTo X
    O = O & M
    If Not IsLasChrUCas(O) Then GoTo X
    GoTo Again
X:
    ShfCml = O
End Function

Function IsLasChrUCas(O$) As Boolean
Dim L$: L = LasChr(O): If L = "" Then Exit Function
IsLasChrUCas = IsAscUCas(Asc(L))
End Function

Function ShfCapCml$(ONm)
Dim P%: P = NxtUcPos(ONm)
If P = 0 Then
    ShfCapCml = ONm
    ONm = ""
Else
    ShfCapCml = Left(ONm, P - 1)
    ONm = Mid(ONm, P)
End If
End Function

Function NxtUcPos%(S)
NxtUcPos = UcPos(S, 2)
End Function

Function UcPos%(S, Fm%)
Dim J%: For J = Fm To Len(S)
    If IsAscUCas(Asc(Mid(S, J, 1))) Then UcPos = J: Exit Function
Next
End Function

Function CmlAyzNy(Ny$()) As String()
Dim I, Nm$
For Each I In Itr(Ny)
    Nm = I
    PushI CmlAyzNy, CmlAy(Nm)
Next
End Function

Function Cmlss(Nm)
':Cmlss: :SS
Cmlss = JnSpc(CmlAy(Nm))
End Function

Function CmlssAy(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushI CmlssAy, Cmlss(N)
Next
End Function

Function CmlAetzNy(Ny$()) As Dictionary
Set CmlAetzNy = Aet(CmlAyzNy(Ny))
End Function

Function DotCml$(Nm)
DotCml = JnDot(CmlAy(Nm))
End Function

Function FstCml$(S)
FstCml = ShfCml(CStr(S))
End Function

Function FstCmlAy(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushI FstCmlAy, FstCml(N)
Next
End Function

Function AscN%(S, N&)
AscN = Asc(Mid(S, N, 1))
End Function

Function IsAscCmlChr(A%) As Boolean
Select Case True
Case IsAscLetter(A), IsAscDig(A), IsAscLDash(A): IsAscCmlChr = True
End Select
End Function

Function IsAscFstCmlChr(A%) As Boolean
If IsAscLDash(A) Then Exit Function
IsAscFstCmlChr = IsAscCmlChr(A)
End Function

Function IsCmlUL(Cml$) As Boolean
Select Case True
Case Len(Cml) <> 2, Not IsAscUCas(FstAsc(Cml)), Not IsAscLCas(SndAsc(Cml))
Case Else: IsCmlUL = LCase(FstChr(Cml)) = SndChr(Cml)
End Select
End Function

Function RmvDigSfx$(S)
Dim J%
For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then RmvDigSfx = Left(S, J): Exit Function
Next
End Function

Function RmvLDashSfx$(S)
Dim J%
For J = Len(S) To 1 Step -1
    If Mid(S, J, 1) <> "_" Then RmvLDashSfx = Left(S, J): Exit Function
Next
End Function

Function CmlRel(Ny$()) As Dictionary
Set CmlRel = Rel(CmlssAy(Ny))
End Function
