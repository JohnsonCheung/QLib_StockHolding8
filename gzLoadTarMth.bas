Attribute VB_Name = "gzLoadTarMth"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzLoadTarMth."
Private Type TarMthDs
    StmTMthDy() As Variant
    BusTMthDy() As Variant
    L1TMthDy() As Variant
    L2TMthDy() As Variant
    L3TMthDy() As Variant
    L4TMthDy() As Variant
    SkuTMthDy() As Variant
    Msg() As String
End Type
Sub LoadTarMth()
Dim A$: A = InputBox("Stocking Report file name:")
If A = "" Then Exit Sub
If Dir(A) = "" Then MsgBox "File not found", vbInformation: Exit Sub
Dim X As New Excel.Application: X.Visible = False
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(A)
If Not IsShWb(Wb) Then
    Dim M$: M = MisShWsMsg(Wb)
    Wb.Close
    MsgBox M, vbCritical
    Wb.Close: Exit Sub
End If
Dim D As TarMthDs: D = TarMthDs(Wb): Wb.Close
BrwEr D.Msg

RplTarMth "Stm", D.StmTMthDy
RplTarMth "Bus", D.BusTMthDy
RplTarMth "L1", D.L1TMthDy
RplTarMth "L2", D.L2TMthDy
RplTarMth "L3", D.L3TMthDy
RplTarMth "L4", D.L4TMthDy
RplTarMth "Sku", D.L4TMthDy
End Sub
Function TarMthDs(Wb As Workbook) As TarMthDs

End Function

Private Sub RplTarMth(PHItm$, PHItmTMthDy())

End Sub
Private Function IsShWb(Wb As Workbook) As Boolean
Dim N$(): N = Wny(Wb)
Dim I: For Each I In PH7Ay
    If Not HasEle(N, "StkHld " & I) Then Exit Function
Next
IsShWb = True
End Function
Private Function MisShWsMsg$(Wb As Workbook)
Dim A$(): A = Wny(Wb)
Dim E$(): E = Split(ExpandPfxSS("StkHld ", PH7ss), " ")
Dim M$(): M = MinusSy(E, A)
If Si(M) = 0 Then Exit Function
MisShWsMsg = "This StkHld Ws is missing:" & vbCrLf & Join(M, vbCrLf)
End Function
