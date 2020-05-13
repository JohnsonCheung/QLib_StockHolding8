Attribute VB_Name = "MxIdeMthOpAdd"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeMthOpAdd."

Sub AddMthByCd(Md As CodeModule, Mthn, CdLines$)
Md.AddFromString NwSubSrcl(Mthn, CdLines)
End Sub

Function NwSubSrcl$(Mthn, CdLines, Optional Mdy$)
NwSubSrcl = AddSfxSpcIf(Mdy) & "Sub ?()" & vbCrLf & "End Sub"
End Function
