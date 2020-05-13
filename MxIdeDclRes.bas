Attribute VB_Name = "MxIdeDclRes"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeRes."
#If ResAA Then
'sdfsdf
#End If

Sub DcrzN__Tst():                                             MsgBox Dcrl(Dcl(Md("MxIdeDclRes")), "ResAA"):               End Sub
Function DcrzN(Resn$):                                DcrzN = DcrzPN(CPj, Resn):                                          End Function
Function DcrzMN(M As CodeModule, Resn$) As String(): DcrzMN = Dcr(Dcl(M), Resn):                                          End Function
Function Dcrl$(Dcl$(), Resn$):                         Dcrl = JnCrLf(Dcr(Dcl, Resn)):                                     End Function
Function Dcr(Dcl$(), Resn$) As String():                Dcr = AmRmvFstChr(AeFstLas(Dcrb(Dcl, Resn))):                     End Function  '#Dcl-Res-Lines#
Private Function Dcrb(Dcl$(), Resn$) As String():      Dcrb = AwBei(Dcl, DcrBei(Dcl, Resn)):                End Function  '#Dcl-Res-block#
Function IsResn$(Nm$):                               IsResn = HasPfx(Nm, "Res"):                                          End Function
Function DcrBix&(Dcl$(), Resn$):                     DcrBix = SrcIx(Dcl, Dcrln(Resn)):                                    End Function
Function DcrEix&(Dcl$(), Fm&):                       DcrEix = SrcIx(Dcl, "#End If", Fm):                                  End Function
Function Dcrln$(Resn$):                               Dcrln = RplQ("#If ? Then", Resn):                                   End Function
Sub ChkIsDcrn(Nm$):                                           ThwTrue IsResn(Nm), CSub, "Given Nm is not Resn", "Nm", Nm: End Sub
Function DcrBei(Dcl$(), Resn$) As Bei
Dim B&: B = DcrBix(Dcl, Resn)
DcrBei = Bei(B, DcrEix(Dcl, B))
End Function

Function DcrzPN(P As VBProject, Resn$) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    DcrzPN = DcrzMN(C.CodeModule, Resn)
    If Si(DcrzPN) > 0 Then Exit Function
Next
End Function

