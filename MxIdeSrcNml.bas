Attribute VB_Name = "MxIdeSrcNml"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcNml."
Function NmlSrc(XSrc$()) As String() ' #Normaliz-XSrc#
'Ret: Normalized-Src from Exported-Src @@
':XSrc: :Src #Exported-Source# ! after :Cmp.Export, the file with have serval lines added.  This :Src is known as :XSrc
NmlSrc = RmvAtrLines(Rmv4ClassLines(XSrc))
End Function
Function RmvAtrLines(Src$()) As String()
Dim Fm%
    Dim J%: For J = 0 To UB(Src)
        If Not HasPfx(Src(J), "Attribute ") Then
            Fm = J
            GoTo X
        End If
    Next
X:
RmvAtrLines = AwBix(Src, Fm)
End Function
Function RmvClassHdrLines(XSrc$()) As String()
RmvClassHdrLines = W1RmvAttrLines(W1Rmv4ClassLines(XSrc))
End Function
Private Function W1Rmv4ClassLines(XSrc$()) As String()
If Si(XSrc) = 0 Then Exit Function
If XSrc(0) = "VERSION 1.0 CLASS" Then
    W1Rmv4ClassLines = AwBix(XSrc, 4)
Else
    W1Rmv4ClassLines = XSrc
End If
End Function
Private Function W1RmvAttrLines(XSrc$()) As String(): W1RmvAttrLines = AwBix(XSrc, W1NonAttrBix(XSrc)): End Function
Private Function W1NonAttrBix&(XSrc$())
Dim J%: For J = 0 To UB(XSrc)
    If Not HasPfx(XSrc(J), "Attribute") Then W1NonAttrBix = J: Exit Function
Next
W1NonAttrBix = Si(XSrc)
End Function

