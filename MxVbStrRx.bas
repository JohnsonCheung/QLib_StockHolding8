Attribute VB_Name = "MxVbStrRx"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Dta"
Const CMod$ = CLib & "MxVbStrRx."
Function SMchsy(S, SPatn$, Optional NSMch% = 1) As String() ' #SubMatch-Str-Array# ret :SMchsy, which SubMch-N-StrAy of @S by @SPatn, where @SPatn is Sub-Patn-Of-N and N = @NSMch
Stop
SMchsy = SMchsyzR(S, Rx(SPatn), NSMch)
End Function
Function SMchsyzR(S, R As RegExp, Optional NSMch% = 1) As String() ' ret :SMchsy, which SubMch-N-StrAy of @S by @R, where @R should has Sub-Patn-Of-N and N = @NSMch
Dim M As Match: For Each M In Mchc(S, R)
    PushS SMchsyzR, WWSMchs(M.SubMatches, NSMch)
Next
End Function

Function SMchs$(S, SPatn$, Optional NSMch% = 1) ' ret fst :SMchs, which SubMch-Str-N of S by @SPatn, where @SPatn is Sub-Patn-Of-N and N = @NSMch
'SPatn:Cml #Sub-Patn#  It is a patn with at least 1 sub-pattern.  Sub-Patn is a patn inside bracket.
'SMchs:: :Mchs #Sub-Match-Str# a submatched substr of @NSMch of @S.
'@SPatn:: :Patn #Sub-Patn# A patn with subPatn, ie there are () inside patn
Dim M As Match: Set M = Mch(S, Rx(SPatn))
If IsNothing(M) Then Exit Function
SMchs = WWSMchs(M.SubMatches, NSMch)
End Function

Function SMchszR$(S, Rx As RegExp)
SMchszR = MchszC(Mchc(S, Rx))
End Function

Function Mchsy(S, Patn$) As String() ' Sy of SubStr of each ele of @Sy by @Patn
Mchsy = MchsyzR(S, Rx(Patn, MultiLine:=True, IsGlobal:=True))
End Function

'--
Private Sub MchsyzR__Tst()
BrwAy MchsyzR(SrclP, HshnRx)
End Sub

Function MchsyzR(S, R As RegExp) As String() ' Sy of (All SubStr of each ele of @Sy by @Rx)
MchsyzR = SyzItv(Mchc(S, R))
End Function

'--
Function Mchs$(S, Patn$) ' ret fst :Mchs by @Patn.
'Mchs:: :Str #Matched-Str# a substr by regexp, which can be :Patn or :Rx
Mchs = MchszR(S, Rx(Patn))
End Function

Function MchszR$(S, Rx As RegExp) ' ret frist :Mchs by @Rx.
MchszR = MchszC(Mchc(S, Rx))
End Function

Function SMchszC$(C As MatchCollection) ' Ret :SMchss MchSubStr of @C
With C
Dim CC As MatchCollection
    If .Count = 0 Then Exit Function
    With CvMch(.Item(0)).SubMatches
        
    End With
        
End With
End Function

Function MchszC$(C As MatchCollection)
With C
    If .Count = 0 Then Exit Function
    MchszC = CvMch(.Item(0)).Value
End With

End Function

Function C12zPatn(Ln, Patn$) As C12
C12zPatn = C12zRx(Ln, Rx(Patn))
End Function

Function C12zRx(S, Rx As RegExp) As C12
Dim M As Match: Set M = Mch(S, Rx)
If IsNothing(M) Then Exit Function
Dim O As C12
With M
    O.C1 = .FirstIndex + 1
    O.C2 = O.C1 + .Length
End With
C12zRx = O
End Function

Function CntzRx&(S, Rx As RegExp)
CntzRx = Mchc(S, Rx).Count + 1
End Function

Function IsMch(S, Rx As RegExp) As Boolean
IsMch = Rx.Test(S)
End Function

'--
Private Sub Mch__Tst()
Dim A As MatchCollection
Dim R  As RegExp: Set R = Rx("m[ae]n")
Set A = Mch("alskdflfmEnsdklf", R)
Stop
End Sub

'--
Function CvRe(A) As RegExp
Set CvRe = A
End Function

Function Rx(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
'Rx:: :RegExp Vb regualr express referring
Const CSub$ = CMod & "Rx"
If Patn = "" Then Thw CSub, "Given @Patn is blank"
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set Rx = O
End Function

Private Sub ReRpl__Tst()
Dim R As RegExp: Set R = Rx("(.+)(m[ae]n)(.+)")
Dim Act$: Act = R.Replace("a men is male", "$1male$3")
Ass Act = "a male is male"
End Sub

Function Mch(S, Rx As RegExp) As Match: Set Mch = FstItm(Rx.Execute(S)): End Function '#Match-Object# Fst Vb RegExp Match Object
Function HasRx(S, Rx As RegExp) As Boolean: HasRx = Rx.Test(S): End Function

Function MchsyzC(C As MatchCollection) As String()
Dim I As Match: For Each I In C
    PushI MchsyzC, I.Value
Next
End Function

Private Function WWSMchs$(S As SubMatches, NSMch%)
With S
    If .Count < NSMch Then Thw CSub, "@SPatn matches @S with subMatch less that Given @NSMch", "NSMch Matched-SubMatches-Count", NSMch
    WWSMchs = .Item(NSMch - 1)
End With
End Function

Private Function Mchc(S, Rx As RegExp) As MatchCollection: Set Mchc = Rx.Execute(S): End Function 'Mchc:: #Match-Collection-Object#

