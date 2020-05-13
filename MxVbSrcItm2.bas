Attribute VB_Name = "MxVbSrcItm2"
Option Compare Text
Option Explicit
Const CNs$ = "Src.Itm"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbSrcItm2."
Private Sub SrcItm__Tst()
MsgBox SrcItm("Private Sub SrcItm")
End Sub

Private Sub IsVbItm__Tst()
MsgBox IsVbItm("Sub")
End Sub

Function IsVbItm(Itm) As Boolean
':SrcItm: :S ! One of :VbItmAy
':VbItmAy: :Ny ! One of {Function Sub Type Enum Property Dim Const}
IsVbItm = HasEle(VbItmAy, Itm)
End Function

Sub ChkIsVbItm(SrcItm, Fun$)
If Not IsVbItm(SrcItm) Then Thw Fun, "@SrcItm should be SrcItm", "@SrcItm Vdt-SrcItm", SrcItm, JnSpc(VbItmAy)
End Sub

Function SrcItm$(Ln)
Dim O$: O = T1(RmvMdy(Ln))
If IsVbItm(O) Then SrcItm = O
End Function

Function VbItmAy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = SyzSS("Function Sub Type Enum Property Dim Const Option Implements")
End If
VbItmAy = Y
End Function

Private Sub VbItmAyzSrc__Tst()
Brw VbItmAyzSrc(SrcV)
End Sub

Function VbItmAyzSrc(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNBNDup VbItmAyzSrc, SrcItm(L)
Next
End Function
