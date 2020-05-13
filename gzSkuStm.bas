Attribute VB_Name = "gzSkuStm"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzSkuStm."
#If Doc Then
':Stm:Cml #Stream# :Chr must be M or U
#End If


Function WrongStmSku(SKU$(), Stm$) As String()
Dim Good$(): Good = SkuyzStm(Stm)
Dim S: For Each S In Itr(SKU)
    If Not HasEle(Good, S) Then
        PushI WrongStmSku, S
    End If
Next
End Function
'-

Function SkuyzStm(Stm$) As String(): SkuyzStm = SyzCQ(SkuSqlzStm(Stm)): End Function
Function MHSkuy() As String(): MHSkuy = SkuyzStm("M"): End Function
Function UDSkuy() As String(): UDSkuy = SkuyzStm("U"): End Function

Private Function SkuSqlzStm$(Stm$): SkuSqlzStm = SkuSql(IIf(Stm = "M", "<>", "=")): End Function
Private Function SkuSqlTp$(): SkuSqlTp = "Select Sku from Sku where Topaz in (Select Topaz From Topaz where Left(CdTopaz,3) ? 'UDV')": End Function
Private Function SkuSql$(Op$): SkuSql = RplQ(SkuSqlTp, Op): End Function

