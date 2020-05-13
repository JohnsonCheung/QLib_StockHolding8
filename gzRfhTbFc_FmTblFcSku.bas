Attribute VB_Name = "gzRfhTbFc_FmTblFcSku"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRfhTbFc_FmTblFcSku."
Private Type FcCoSum
    NFc As Long
    NSku As Long
    SC   As Double
End Type
Private Type FcSum
    HK As FcCoSum
    MO As FcCoSum
    Siz As Long
    Tim As Date
End Type

Private Sub RfhTbFc_Fm_TblFcSku__Tst()
RfhTbFc_Fm_TblFcSku LasUDFc
End Sub

Sub RfhTbFc_Fm_TblFcSku(A As StmYM)
With FcSum(A)
RunCQ "Update Fc set" & _
" DteLoad=Now," & _
" C86NFc=" & .HK.NFc & "," & _
" C86NSku=" & .HK.NSku & "," & _
" C86Sc=" & .HK.SC & "," & _
" C87NFc=" & .MO.NFc & "," & _
" C87NSku=" & .MO.NSku & "," & _
" C87Sc=" & .MO.SC & _
WhereFcStm(A)
End With
End Sub

Private Function FcSum(A As StmYM) As FcSum
Dim W$: W = WhereFcStm(A)
Dim Sql$: Sql = "Select Distinct Co,Count(*) As NSku," & _
"Sum(Nz(M01,0)+Nz(M02,0)+Nz(M03,0)+Nz(M04,0)+Nz(M05,0)" & _
   "+Nz(M06,0)+Nz(M07,0)+Nz(M08,0)+Nz(M09,0)+Nz(M10,0)" & _
   "+Nz(M11,0)+Nz(M12,0)+Nz(M13,0)+Nz(M14,0)+Nz(M15,0)) As SC," & _
"Sum(" & _
"IIf(Nz(M01,0)=0,0,1)+IIf(Nz(M02,0)=0,0,1)+IIf(Nz(M03,0)=0,0,1)" & _
"+IIf(Nz(M04,0)=0,0,1)+IIf(Nz(M05,0)=0,0,1)+IIf(Nz(M06,0)=0,0,1)" & _
"+IIf(Nz(M07,0)=0,0,1)+IIf(Nz(M08,0)=0,0,1)+IIf(Nz(M09,0)=0,0,1)" & _
"+IIf(Nz(M10,0)=0,0,1)+IIf(Nz(M11,0)=0,0,1)+IIf(Nz(M12,0)=0,0,1)" & _
"+IIf(Nz(M13,0)=0,0,1)+IIf(Nz(M14,0)=0,0,1)+IIf(Nz(M15,0)=0,0,1)) As NFc" & _
" from FcSku" & W & QpGp_FF("Co")
Dim Rs As DAO.Recordset: Set Rs = CurrentDb.OpenRecordset(Sql)
With Rs
    While Not .EOF
        Dim Co As Byte: Co = !Co
        Select Case Co
        Case 86: FcSum.HK = FcCoSumzRs(Rs)
        Case 87: FcSum.MO = FcCoSumzRs(Rs)
        Case Else: RaiseCo Co
        End Select
        .MoveNext
    Wend
End With
End Function

Private Function FcCoSumzRs(Rs As Recordset) As FcCoSum
Dim O As FcCoSum
With Rs
    O.NSku = !NSku
    O.SC = !SC
    O.NFc = !NFc
End With
FcCoSumzRs = O
End Function
