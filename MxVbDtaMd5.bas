Attribute VB_Name = "MxVbDtaMd5"
Option Compare Text
Option Explicit
Const CNs$ = "Md5"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbDtaMd5."
Private enc_

Private Property Get enc() As Object
Static X As Object, Y As Boolean
If Not Y Then Y = True: Set X = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
Set enc = X
End Property

Function BytAyHex$(A() As Byte)
Dim O$()
Dim I:  For Each I In Itr(A)
    PushI O, Right("0" & Hex(I), 2)
Next
BytAyHex = Join(O, "")
End Function

Function MD5$(S$)
Dim textBytes() As Byte: textBytes = S
MD5 = FmtBytAy(CvBytAy(enc.ComputeHash_2((textBytes))))
End Function

Function FmtBytAy$(A() As Byte, Optional NBytTogether% = 2)
Dim O$(), N%
N = Si(A)
If N Mod NBytTogether <> 0 Then Thw CSub, "Si-of-@BytAy is not multiple of @NBytTogether", "Si-BytAy NBytTogether", N, NBytTogether
Dim J%: For J = 1 To N \ NBytTogether
    PushI O, BytAyHex(WhBytAy(A, J, NBytTogether))
Next
FmtBytAy = JnSpc(O)
End Function

Function WhBytAy(A() As Byte, IthBlk%, BlkSi%) As Byte()
Dim Offset%: Offset = (IthBlk - 1) * BlkSi
Dim J%: For J = 0 To BlkSi - 1
    PushI WhBytAy, A(Offset + J)
Next
End Function
