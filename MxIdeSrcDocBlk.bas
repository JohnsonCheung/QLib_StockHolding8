Attribute VB_Name = "MxIdeSrcDocBlk"
Option Explicit
Option Compare Text
#If Doc Then
#End If
'**OptSrc
Private Sub OptSrc__Tst()
Dim O() As S12
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushS12 O, S12(C.Name, DocBlkl(Dcl(C.CodeModule)))
Next

BrwS12y RmvBlnkS2(O)
End Sub
Function OptSrc(Src$(), Optn$) As String():       OptSrc = AwBei(Src, OptSrcBei(Src, Optn)): End Function
Function RmvOptSrc(Src$(), Optn$) As String(): RmvOptSrc = AeBei(Src, OptSrcBei(Src, Optn)): End Function
Function OptSrcBei(Src$(), Optn$) As Bei
Dim B%: B = Bix(Src, Optn)
OptSrcBei = Bei(B, Eix(Src, B))
End Function
Private Function Bix%(Src$(), Optn$): Bix = SrcIx(Src, Ln(Optn)):           End Function
Private Function Eix%(Src$(), Bix%):  Eix = SrcIx(Src, "#End If", Bix + 1): End Function
Private Function Ln$(Optn$):           Ln = FmtQQ("#If ? Then", Optn):      End Function

'**DocBlk
Function DocBlkl$(Dcl$()):                DocBlkl = JnCrLf(DocBlk(Dcl)):               End Function
Function DocBlk(Dcl$()) As String():       DocBlk = AmRmvFstChr(OptSrc(Dcl, "Doc")):   End Function
Function RmvDocBlk(Src$()) As String(): RmvDocBlk = AeBei(Src, OptSrcBei(Src, "Doc")): End Function
