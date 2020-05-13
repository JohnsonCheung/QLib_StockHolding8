Attribute VB_Name = "MxVbStrRpl"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrRpl."
':Q: :S #Str-With-QuestionMark#
Private Sub RplBet__Tst()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub

Private Sub RplPfx__Tst()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub

Function RplCr$(S)
RplCr = Replace(S, vbCr, " ")
End Function
Function RplCrLf$(S)
RplCrLf = RplLf(RplCr(S))
End Function
Function RplLf$(S)
RplLf = Replace(S, vbLf, " ")
End Function
Function RplVbl$(S)
RplVbl = RplVBar(S)
End Function
Function RplVBar$(S)
RplVBar = Replace(S, "|", vbCrLf)
End Function
Function RplBet$(S, By$, S1$, S2$)
Dim P1%, P2%, B$, C$
P1 = InStr(S, S1)
If P1 = 0 Then Stop
P2 = InStr(P1 + Len(S1), S, S2)
If P2 = 0 Then Stop
B = Left(S, P1 + Len(S1) - 1)
C = Mid(S, P2 + Len(S2) - 1)
RplBet = B & By & C
End Function

Function Rpl2DblQ$(S)
'Ret :S #Rpl-2DblQ-To-Blnk#
Rpl2DblQ = Replace(S, vb2DblQ, "")
End Function

Function RplDblSpc$(S)
Dim O$: O = Trim(S)
Dim J&
While HasSubStr(O, "  ")
    J = J + 1: If J > 10000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function RplFstChr$(S, By$)
RplFstChr = By & RmvFstChr(S)
End Function
Function DisChry(S) As String() ' ret Dis-Chry from S
Dim O$(), J&: For J = 1 To Len(S): PushI O, Mid(S, J, 1): Next
DisChry = AwDis(O)
End Function
Function RplPfx(S, Fm$, ToPfx$)
If HasPfx(S, Fm) Then
    RplPfx = ToPfx & RmvPfx(S, Fm)
Else
    RplPfx = S
End If
End Function

Sub PurePun__Tst()
VcAy PurePun(SrclP)
End Sub

Function PurePun(S) As String()
PurePun = QSrt(AwNB(DisChry(RmvDblSpc(RplCrLf(RplAlpNum(S))))))
End Function

Function AlpNumRx() As RegExp
Static X As RegExp
If IsNothing(X) Then Set X = Rx("[0-9a-zA-Z]", IsGlobal:=True)
Set AlpNumRx = X
End Function

Function PunRx() As RegExp
Static X As RegExp
If IsNothing(X) Then Set X = Rx("[!""#$%&'()*+,-/:;<=>?@[\\\]^_`{\|}~]", IsGlobal:=True)
Set PunRx = X
'  0 1 2 3 4 5 6 7 8 9 A B C D E F
'0                
'1                
'2   ! " # $ % & ' ( ) * + , - . /
'3 0 1 2 3 4 5 6 7 8 9 : ; < = > ?
'4 @ A B C D E F G H I J K L M N O
'5 P Q R S T U V W X Y Z [ \ ] ^ _
'6 ` a b c d e f g h i j k l m n o
'7 p q r s t u v w x y z { | } ~ 
End Function
Private Sub RplPun__Tst()
VcStr RplPun(SrclP)
End Sub

Function RplPun$(S)
RplPun = PunRx.Replace(S, " ")
End Function

Function RplAlpNum$(S)
RplAlpNum = AlpNumRx.Replace(S, " ")
End Function

Function RplQ$(Q, By)
RplQ = Replace(Q, "?", By)
End Function

Function DyoAyAv(AyAv()) As Variant()
If Si(AyAv) = 0 Then Exit Function
Dim UAy%: UAy = UB(AyAv)
Dim URec&: URec = UB(AyAv(0))
Dim ODy(): ReDim ODy(URec)
Dim R&: For R = 0 To URec
    Dim Dr(): ReDim Dr(UAy)
    Dim C%: For C = 0 To UAy
        Dr(C) = AyAv(C)(R)
    Next
    ODy(R) = Dr
Next
DyoAyAv = ODy
End Function
Function SyzMacro(RplMacro$, ParamArray ByAyAp()) As String()
Dim AyAv(): AyAv = ByAyAp
SyzMacro = SyzMacroDy(RplMacro, DyoAyAv(AyAv))
End Function

Function SyzMacroDy(RplMacro$, ByDy()) As String()
Const CSub$ = CMod & "SyzMacroDy"
If Si(ByDy) = 0 Then Exit Function
Dim M$():     M = MacroNy(RplMacro, InlBkt:=True)
Dim URec&: URec = UB(ByDy)
Dim UFld%: UFld = UB(ByDy(0))

If UB(M) <> UFld Then Thw CSub, "UFld should = UB(MacroNy)", "UFld UB(MacroNy)", UFld, UB(M)

'-- O --
Dim O$(): ReDim O(URec)
Dim J&, Dr: For Each Dr In ByDy
    O(J) = SzMacro(RplMacro, M, Dr)
    J = J + 1
Next
SyzMacroDy = O
End Function

Function SzMacro$(RplMacro$, OfMacroNy$(), ByDr)
Dim O$: O = RplMacro
Dim V, J%: For Each V In ByDr
    O = Replace(O, OfMacroNy(J), V)
    J = J + 1
Next
SzMacro = O
End Function

Function SyzQAy(Q, ByAy) As String()
Dim By: For Each By In Itr(ByAy)
    PushI SyzQAy, RplQ(Q, By)
Next
End Function

Private Sub RplBet3__Tst()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub
