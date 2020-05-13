Attribute VB_Name = "MxXlsReadFx"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsReadFx."
Type FxwcSy
    Fx As String
    W As String
    C As String
    Sy() As String
End Type

Function FxwcSy(Fx$, W$, C$, Sy$()) As FxwcSy
With FxwcSy
    .Fx = Fx
    .W = W
    .C = C
    .Sy = Sy
End With
End Function

Function SyzFxQ(Fx$, Q$) As String()
SyzFxQ = SyzArs(ArszFxQ(Fx, Q))
End Function

Function DisSyzFxwc(Fx$, W$, C$) As String()
Dim Q$: Q = SqlSel_F_T(C, AxTbn(W))
DisSyzFxwc = SyzArs(ArszFxQ(Fx, Q))
End Function

Function VyzFxQ(Fx$, Q$) As Variant()
VyzFxQ = ColzArs(ArszFxQ(Fx, Q))
End Function

Function DrszFxw(Fx$, W) As Drs
DrszFxw = DrszArs(ArszFxw(Fx, W))
End Function

Function FldIsBlnkBexp$(F)
':Bexp: :S ! #Bool-Epr-Str#
FldIsBlnkBexp = FmtQQ("Trim(Nz([?],''))=''", F)
End Function

Function DtzFxw(Fx$, Optional Wsn0$) As Dt
Dim W$: W = DftWsn(Wsn0, Fx)
DtzFxw = DtzDrs(DrszFxw(Fx, W), W)
End Function

Function IntColzFx(Fx$, W$, C$) As Integer():      IntColzFx = IntoColzFx(EmpIntAy, Fx, W, C): End Function
Function StrColzFx(Fx$, W$, C$) As String():       StrColzFx = IntoColzFx(EmpSy, Fx, W, C): End Function
Function ColzFx(Fx$, W$, C$) As Variant():            ColzFx = IntoColzFx(EmpAv, Fx, W, C): End Function ' :Av '#Col-Value-Ay#
Private Function IntoColzFx(IntoAy, Fx$, W$, C$): IntoColzFx = IntoColzArs(IntoAy, ArszFx(Fx, W, C)): End Function
Function IntoColzArs(IntoAy, A As ADODB.Recordset, Optional C = 0)
Dim O: O = IntoAy: Erase O
With A
    While Not .EOF
        PushI O, Nz(.Fields(C).Value, Empty)
        .MoveNext
    Wend
    .Close
End With
IntoColzArs = O
End Function

Function ArszFxDis(Fx$, W$, DisC) As ADODB.Recordset
Set ArszFxDis = ArszCnq(CnzFx(Fx), SqlSel_F_T(DisC, AxTbn(W)))
End Function

Function DisFvyzFx(Fx$, W$, C$) As Variant()
':DisFvy :Av #QpDis-Fld-Vy#
DisFvyzFx = IntoDisFvyzFx(EmpAv, Fx, W, C)
End Function

Function DisFsyzFx(Fx$, W$, C$) As String()
'DisFsy: :Sy #QpDis-Fld-Sy#
DisFsyzFx = IntoDisFvyzFx(EmpSy, Fx, W, C)
End Function

Function ReadFxwcSy(Fx$, W$, C$) As FxwcSy
Dim Sy$(): Sy = DisFsyzFx(Fx, W, C)
ReadFxwcSy = FxwcSy(Fx, W, C, Sy)
End Function

Private Function IntoDisFvyzFx(IntoAy, Fx$, W$, C$)
IntoDisFvyzFx = IntoColzArs(IntoAy, ArszFxDis(Fx, W, C))
End Function

