Attribute VB_Name = "MxDaoDbSchmUd"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbSchmUd."
Public Const StdEleLines$ = _
"E Crt Dte;Req;Dft=Now" & vbCrLf & _
"E Tim Dte" & vbCrLf & _
"E Lng Lng" & vbCrLf & _
"E Mem Mem" & vbCrLf & _
"E Dte Dte" & vbCrLf & _
"E Nm  Txt;Req;Sz=50"
Public Const StdETFLines$ = _
"ETF Nm  * *Nm          " & vbCrLf & _
"ETF Tim * *Tim         " & vbCrLf & _
"ETF Dte * *Dte         " & vbCrLf & _
"ETF Crt * CrtTim       " & vbCrLf & _
"ETF Lng * Si           " & vbCrLf & _
"ETF Mem * Lines *Ft *Fx"
'== Sms SchmSrc =================================================
Type SmsTbl:       Lno As Integer: Tbn As String:  Fny() As String: SkFny() As String:            End Type 'Deriving(Ctor Ay)
Type SmsEleFld:    Lno As Integer: Elen As String: FldLikAy() As String:                          End Type 'Deriving(Ctor Ay)
Type SmsEle:       Lno As Integer: Elen As String: EleStr As String:                              End Type 'Deriving(Ctor Ay)
Type SmsKey:       Lno As Integer: Tbn As String:  Keyn As String: IsUniq As Boolean: Fny() As String: End Type 'Deriving(Ctor Ay)
Type SmsTblDes:    Lno As Integer: Tbn As String:  Des As String:                                 End Type 'Deriving(Ctor Ay)
Type SmsTblFldDes: Lno As Integer: Tbn As String:  Fldn As String: Des As String:                 End Type 'Deriving(Ctor Ay)
Type SmsFldDes:    Lno As Integer: Fldn As String: Des As String: End Type 'Deriving(Ctor Ay)
Type SchmSrc
    Tbl() As SmsTbl
    EleFld() As SmsEleFld
    Ele() As SmsEle
    TblDes() As SmsTblDes
    TblFldDes() As SmsTblFldDes
    FldDes() As SmsFldDes
    Key() As SmsKey
End Type 'Deriving(Ctor)
'==Smd SchmDta ===================================================
Type smdTbl:    Tbn As String: Fny() As String:        End Type 'Deriving(Ctor Ay)
Type SmdSk:     Tbn As String: SkFny() As String:      End Type 'Deriving(Ctor Ay)
Type smdKey:    Tbn As String: KeyFny() As String:     End Type 'Deriving(Ctor Ay)
Type smdFk:     Tbn As String: ParTbn As String:               End Type 'Deriving(Ctor Ay)
Type smdTblDes: Tbn As String: Des As String:                  End Type 'Deriving(Ctor Ay)
Type smdTFDes:  Tbn As String: Fldn As String: Des As String:  End Type 'Deriving(Ctor Ay)
Type smdFldDes: Fldn As String: Des As String:                  End Type 'Deriving(Ctor Ay)
Type SchmDta
    Tbl() As smdTbl
    PkTny() As String
    Sk() As SmdSk
    Key() As smdKey
    Fk() As smdFk
    TblDes() As smdTblDes
    TFDes() As smdTFDes
    FldDes() As smdFldDes
End Type
'==Smd SchmEr ===================================================
Enum eSmeTbl: eSmeT: End Enum
Enum eSmeEleF: eSmeEF: End Enum
Enum eSmeEle: eSmeE: End Enum
Enum eSmeKey: eSmeK: End Enum
Enum eSmeTblDes: eSmeTDes: End Enum
Enum eSmeTFDes: eSmeTFDes_: End Enum
Enum eSmeFldDes: eSmeFDes: End Enum

Type smeTbl:  Lno As Integer: T As eSmeTbl: End Type 'Deriving(Ctor Ay)
Type smeEleF: Lno As Integer: T As eSmeEleF: End Type 'Deriving(Ctor Ay)
Type smeEle:  Lno As Integer: T As eSmeEle: End Type 'Deriving(Ctor Ay)
Type smeKey:  Lno As Integer: T As eSmeKey: End Type 'Deriving(Ctor Ay)
Type smeTblDes: Lno As Integer: T As eSmeTblDes: End Type 'Deriving(Ctor Ay)
Type smeTFDes:  Lno As Integer: T As eSmeTFDes: End Type 'Deriving(Ctor Ay)
Type smeFldDes: Lno As Integer: T As eSmeFldDes: End Type 'Deriving(Ctor Ay)
Type SchmEr
    Tbl() As smeTbl
    EleF() As smeEleF
    Ele() As smeEle
    Key() As smeKey
    TblDes() As smeTblDes
    TFDes() As smeTblDes
    FldDes() As smeTblDes
End Type
'==Smp SchmParse =================================================
Type SchmPsr: Dta As SchmDta: Er As SchmEr: End Type 'Deriving(Ctor)
'==Smb SchmBld ===================================================
Type SchmBld
    TdAy() As DAO.TableDef
    PkSqy() As String
    SkSqy() As String
    KeySqy() As String
    FkSqy() As String
    TblDesDi As Dictionary
    FldDesDi As Dictionary
End Type
Function SmsTbl(Lno, Tbn, Fny$(), SkFny$()) As SmsTbl
With SmsTbl
    .Lno = Lno
    .Tbn = Tbn
    .Fny = Fny
    .SkFny = SkFny
End With
End Function
Function AddSmsTbl(A As SmsTbl, B As SmsTbl) As SmsTbl(): PushSmsTbl AddSmsTbl, A: PushSmsTbl AddSmsTbl, B: End Function
Sub PushSmsTblAy(O() As SmsTbl, A() As SmsTbl): Dim J&: For J = 0 To SmsTblUB(A): PushSmsTbl O, A(J): Next: End Sub
Sub PushSmsTbl(O() As SmsTbl, M As SmsTbl): Dim N&: N = SmsTblUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmsTblSi&(A() As SmsTbl): On Error Resume Next: SmsTblSi = UBound(A) + 1: End Function
Function SmsTblUB&(A() As SmsTbl): SmsTblUB = SmsTblSi(A) - 1: End Function
Function SmsEleFld(Lno, Elen, FldLikAy$()) As SmsEleFld
With SmsEleFld
    .Lno = Lno
    .Elen = Elen
    .FldLikAy = FldLikAy
End With
End Function
Function AddSmsEleFld(A As SmsEleFld, B As SmsEleFld) As SmsEleFld(): PushSmsEleFld AddSmsEleFld, A: PushSmsEleFld AddSmsEleFld, B: End Function
Sub PushSmsEleFldAy(O() As SmsEleFld, A() As SmsEleFld): Dim J&: For J = 0 To SmsEleFldUB(A): PushSmsEleFld O, A(J): Next: End Sub
Sub PushSmsEleFld(O() As SmsEleFld, M As SmsEleFld): Dim N&: N = SmsEleFldUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmsEleFldSi&(A() As SmsEleFld): On Error Resume Next: SmsEleFldSi = UBound(A) + 1: End Function
Function SmsEleFldUB&(A() As SmsEleFld): SmsEleFldUB = SmsEleFldSi(A) - 1: End Function
Function SmsEle(Lno, Elen, EleStr) As SmsEle
With SmsEle
    .Lno = Lno
    .Elen = Elen
    .EleStr = EleStr
End With
End Function
Function AddSmsEle(A As SmsEle, B As SmsEle) As SmsEle(): PushSmsEle AddSmsEle, A: PushSmsEle AddSmsEle, B: End Function
Sub PushSmsEleAy(O() As SmsEle, A() As SmsEle): Dim J&: For J = 0 To SmsEleUB(A): PushSmsEle O, A(J): Next: End Sub
Sub PushSmsEle(O() As SmsEle, M As SmsEle): Dim N&: N = SmsEleUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmsEleSi&(A() As SmsEle): On Error Resume Next: SmsEleSi = UBound(A) + 1: End Function
Function SmsEleUB&(A() As SmsEle): SmsEleUB = SmsEleSi(A) - 1: End Function
Function SmsKey(Lno, Tbn, Keyn, IsUniq, Fny$()) As SmsKey
With SmsKey
    .Lno = Lno
    .Tbn = Tbn
    .Keyn = Keyn
    .IsUniq = IsUniq
    .Fny = Fny
End With
End Function
Function AddSmsKey(A As SmsKey, B As SmsKey) As SmsKey(): PushSmsKey AddSmsKey, A: PushSmsKey AddSmsKey, B: End Function
Sub PushSmsKeyAy(O() As SmsKey, A() As SmsKey): Dim J&: For J = 0 To SmsKeyUB(A): PushSmsKey O, A(J): Next: End Sub
Sub PushSmsKey(O() As SmsKey, M As SmsKey): Dim N&: N = SmsKeyUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmsKeySi&(A() As SmsKey): On Error Resume Next: SmsKeySi = UBound(A) + 1: End Function
Function SmsKeyUB&(A() As SmsKey): SmsKeyUB = SmsKeySi(A) - 1: End Function
Function SmsTblDes(Lno, Tbn, Des) As SmsTblDes
With SmsTblDes
    .Lno = Lno
    .Tbn = Tbn
    .Des = Des
End With
End Function
Function AddSmsTblDes(A As SmsTblDes, B As SmsTblDes) As SmsTblDes(): PushSmsTblDes AddSmsTblDes, A: PushSmsTblDes AddSmsTblDes, B: End Function
Sub PushSmsTblDesAy(O() As SmsTblDes, A() As SmsTblDes): Dim J&: For J = 0 To SmsTblDesUB(A): PushSmsTblDes O, A(J): Next: End Sub
Sub PushSmsTblDes(O() As SmsTblDes, M As SmsTblDes): Dim N&: N = SmsTblDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmsTblDesSi&(A() As SmsTblDes): On Error Resume Next: SmsTblDesSi = UBound(A) + 1: End Function
Function SmsTblDesUB&(A() As SmsTblDes): SmsTblDesUB = SmsTblDesSi(A) - 1: End Function
Function SmsTblFldDes(Lno, Tbn, Fldn, Des) As SmsTblFldDes
With SmsTblFldDes
    .Lno = Lno
    .Tbn = Tbn
    .Fldn = Fldn
    .Des = Des
End With
End Function
Function AddSmsTblFldDes(A As SmsTblFldDes, B As SmsTblFldDes) As SmsTblFldDes(): PushSmsTblFldDes AddSmsTblFldDes, A: PushSmsTblFldDes AddSmsTblFldDes, B: End Function
Sub PushSmsTblFldDesAy(O() As SmsTblFldDes, A() As SmsTblFldDes): Dim J&: For J = 0 To SmsTblFldDesUB(A): PushSmsTblFldDes O, A(J): Next: End Sub
Sub PushSmsTblFldDes(O() As SmsTblFldDes, M As SmsTblFldDes): Dim N&: N = SmsTblFldDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmsTblFldDesSi&(A() As SmsTblFldDes): On Error Resume Next: SmsTblFldDesSi = UBound(A) + 1: End Function
Function SmsTblFldDesUB&(A() As SmsTblFldDes): SmsTblFldDesUB = SmsTblFldDesSi(A) - 1: End Function
Function SmsFldDes(Lno, Fldn, Des) As SmsFldDes
With SmsFldDes
    .Lno = Lno
    .Fldn = Fldn
    .Des = Des
End With
End Function
Function AddSmsFldDes(A As SmsFldDes, B As SmsFldDes) As SmsFldDes(): PushSmsFldDes AddSmsFldDes, A: PushSmsFldDes AddSmsFldDes, B: End Function
Sub PushSmsFldDesAy(O() As SmsFldDes, A() As SmsFldDes): Dim J&: For J = 0 To SmsFldDesUB(A): PushSmsFldDes O, A(J): Next: End Sub
Sub PushSmsFldDes(O() As SmsFldDes, M As SmsFldDes): Dim N&: N = SmsFldDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmsFldDesSi&(A() As SmsFldDes): On Error Resume Next: SmsFldDesSi = UBound(A) + 1: End Function
Function SmsFldDesUB&(A() As SmsFldDes): SmsFldDesUB = SmsFldDesSi(A) - 1: End Function
Function SchmSrc(Tbl() As SmsTbl, EleFld() As SmsEleFld, Ele() As SmsEle, TblDes() As SmsTblDes, TblFldDes() As SmsTblFldDes, FldDes() As SmsFldDes, Key() As SmsKey) As SchmSrc
With SchmSrc
    .Tbl = Tbl
    .EleFld = EleFld
    .Ele = Ele
    .TblDes = TblDes
    .TblFldDes = TblFldDes
    .FldDes = FldDes
    .Key = Key
End With
End Function
Function smdTbl(Tbn, Fny$()) As smdTbl
With smdTbl
    .Tbn = Tbn
    .Fny = Fny
End With
End Function
Function AddsmdTbl(A As smdTbl, B As smdTbl) As smdTbl(): PushsmdTbl AddsmdTbl, A: PushsmdTbl AddsmdTbl, B: End Function
Sub PushsmdTblAy(O() As smdTbl, A() As smdTbl): Dim J&: For J = 0 To smdTblUB(A): PushsmdTbl O, A(J): Next: End Sub
Sub PushsmdTbl(O() As smdTbl, M As smdTbl): Dim N&: N = smdTblUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smdTblSi&(A() As smdTbl): On Error Resume Next: smdTblSi = UBound(A) + 1: End Function
Function smdTblUB&(A() As smdTbl): smdTblUB = smdTblSi(A) - 1: End Function
Function SmdSk(Tbn, SkFny$()) As SmdSk
With SmdSk
    .Tbn = Tbn
    .SkFny = SkFny
End With
End Function
Function AddSmdSk(A As SmdSk, B As SmdSk) As SmdSk(): PushSmdSk AddSmdSk, A: PushSmdSk AddSmdSk, B: End Function
Sub PushSmdSkAy(O() As SmdSk, A() As SmdSk): Dim J&: For J = 0 To SmdSkUB(A): PushSmdSk O, A(J): Next: End Sub
Sub PushSmdSk(O() As SmdSk, M As SmdSk): Dim N&: N = SmdSkUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SmdSkSi&(A() As SmdSk): On Error Resume Next: SmdSkSi = UBound(A) + 1: End Function
Function SmdSkUB&(A() As SmdSk): SmdSkUB = SmdSkSi(A) - 1: End Function
Function smdKey(Tbn, KeyFny$()) As smdKey
With smdKey
    .Tbn = Tbn
    .KeyFny = KeyFny
End With
End Function
Function AddsmdKey(A As smdKey, B As smdKey) As smdKey(): PushsmdKey AddsmdKey, A: PushsmdKey AddsmdKey, B: End Function
Sub PushsmdKeyAy(O() As smdKey, A() As smdKey): Dim J&: For J = 0 To smdKeyUB(A): PushsmdKey O, A(J): Next: End Sub
Sub PushsmdKey(O() As smdKey, M As smdKey): Dim N&: N = smdKeyUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smdKeySi&(A() As smdKey): On Error Resume Next: smdKeySi = UBound(A) + 1: End Function
Function smdKeyUB&(A() As smdKey): smdKeyUB = smdKeySi(A) - 1: End Function
Function smdFk(Tbn, ParTbn) As smdFk
With smdFk
    .Tbn = Tbn
    .ParTbn = ParTbn
End With
End Function
Function AddsmdFk(A As smdFk, B As smdFk) As smdFk(): PushsmdFk AddsmdFk, A: PushsmdFk AddsmdFk, B: End Function
Sub PushsmdFkAy(O() As smdFk, A() As smdFk): Dim J&: For J = 0 To smdFkUB(A): PushsmdFk O, A(J): Next: End Sub
Sub PushsmdFk(O() As smdFk, M As smdFk): Dim N&: N = smdFkUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smdFkSi&(A() As smdFk): On Error Resume Next: smdFkSi = UBound(A) + 1: End Function
Function smdFkUB&(A() As smdFk): smdFkUB = smdFkSi(A) - 1: End Function
Function smdTblDes(Tbn, Des) As smdTblDes
With smdTblDes
    .Tbn = Tbn
    .Des = Des
End With
End Function
Function AddsmdTblDes(A As smdTblDes, B As smdTblDes) As smdTblDes(): PushsmdTblDes AddsmdTblDes, A: PushsmdTblDes AddsmdTblDes, B: End Function
Sub PushsmdTblDesAy(O() As smdTblDes, A() As smdTblDes): Dim J&: For J = 0 To smdTblDesUB(A): PushsmdTblDes O, A(J): Next: End Sub
Sub PushsmdTblDes(O() As smdTblDes, M As smdTblDes): Dim N&: N = smdTblDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smdTblDesSi&(A() As smdTblDes): On Error Resume Next: smdTblDesSi = UBound(A) + 1: End Function
Function smdTblDesUB&(A() As smdTblDes): smdTblDesUB = smdTblDesSi(A) - 1: End Function
Function smdTFDes(Tbn, Fldn, Des) As smdTFDes
With smdTFDes
    .Tbn = Tbn
    .Fldn = Fldn
    .Des = Des
End With
End Function
Function AddsmdTFDes(A As smdTFDes, B As smdTFDes) As smdTFDes(): PushsmdTFDes AddsmdTFDes, A: PushsmdTFDes AddsmdTFDes, B: End Function
Sub PushsmdTFDesAy(O() As smdTFDes, A() As smdTFDes): Dim J&: For J = 0 To smdTFDesUB(A): PushsmdTFDes O, A(J): Next: End Sub
Sub PushsmdTFDes(O() As smdTFDes, M As smdTFDes): Dim N&: N = smdTFDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smdTFDesSi&(A() As smdTFDes): On Error Resume Next: smdTFDesSi = UBound(A) + 1: End Function
Function smdTFDesUB&(A() As smdTFDes): smdTFDesUB = smdTFDesSi(A) - 1: End Function
Function smdFldDes(Fldn, Des) As smdFldDes
With smdFldDes
    .Fldn = Fldn
    .Des = Des
End With
End Function
Function AddsmdFldDes(A As smdFldDes, B As smdFldDes) As smdFldDes(): PushsmdFldDes AddsmdFldDes, A: PushsmdFldDes AddsmdFldDes, B: End Function
Sub PushsmdFldDesAy(O() As smdFldDes, A() As smdFldDes): Dim J&: For J = 0 To smdFldDesUB(A): PushsmdFldDes O, A(J): Next: End Sub
Sub PushsmdFldDes(O() As smdFldDes, M As smdFldDes): Dim N&: N = smdFldDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smdFldDesSi&(A() As smdFldDes): On Error Resume Next: smdFldDesSi = UBound(A) + 1: End Function
Function smdFldDesUB&(A() As smdFldDes): smdFldDesUB = smdFldDesSi(A) - 1: End Function
Function smeTbl(Lno, T As eSmeTbl) As smeTbl
With smeTbl
    .Lno = Lno
    .T = T
End With
End Function
Function AddsmeTbl(A As smeTbl, B As smeTbl) As smeTbl(): PushsmeTbl AddsmeTbl, A: PushsmeTbl AddsmeTbl, B: End Function
Sub PushsmeTblAy(O() As smeTbl, A() As smeTbl): Dim J&: For J = 0 To smeTblUB(A): PushsmeTbl O, A(J): Next: End Sub
Sub PushsmeTbl(O() As smeTbl, M As smeTbl): Dim N&: N = smeTblUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smeTblSi&(A() As smeTbl): On Error Resume Next: smeTblSi = UBound(A) + 1: End Function
Function smeTblUB&(A() As smeTbl): smeTblUB = smeTblSi(A) - 1: End Function
Function smeEleF(Lno, T As eSmeEleF) As smeEleF
With smeEleF
    .Lno = Lno
    .T = T
End With
End Function
Function AddsmeEleF(A As smeEleF, B As smeEleF) As smeEleF(): PushsmeEleF AddsmeEleF, A: PushsmeEleF AddsmeEleF, B: End Function
Sub PushsmeEleFAy(O() As smeEleF, A() As smeEleF): Dim J&: For J = 0 To smeEleFUB(A): PushsmeEleF O, A(J): Next: End Sub
Sub PushsmeEleF(O() As smeEleF, M As smeEleF): Dim N&: N = smeEleFUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smeEleFSi&(A() As smeEleF): On Error Resume Next: smeEleFSi = UBound(A) + 1: End Function
Function smeEleFUB&(A() As smeEleF): smeEleFUB = smeEleFSi(A) - 1: End Function
Function smeEle(Lno, T As eSmeEle) As smeEle
With smeEle
    .Lno = Lno
    .T = T
End With
End Function
Function AddsmeEle(A As smeEle, B As smeEle) As smeEle(): PushsmeEle AddsmeEle, A: PushsmeEle AddsmeEle, B: End Function
Sub PushsmeEleAy(O() As smeEle, A() As smeEle): Dim J&: For J = 0 To smeEleUB(A): PushsmeEle O, A(J): Next: End Sub
Sub PushsmeEle(O() As smeEle, M As smeEle): Dim N&: N = smeEleUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smeEleSi&(A() As smeEle): On Error Resume Next: smeEleSi = UBound(A) + 1: End Function
Function smeEleUB&(A() As smeEle): smeEleUB = smeEleSi(A) - 1: End Function
Function smeKey(Lno, T As eSmeKey) As smeKey
With smeKey
    .Lno = Lno
    .T = T
End With
End Function
Function AddsmeKey(A As smeKey, B As smeKey) As smeKey(): PushsmeKey AddsmeKey, A: PushsmeKey AddsmeKey, B: End Function
Sub PushsmeKeyAy(O() As smeKey, A() As smeKey): Dim J&: For J = 0 To smeKeyUB(A): PushsmeKey O, A(J): Next: End Sub
Sub PushsmeKey(O() As smeKey, M As smeKey): Dim N&: N = smeKeyUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smeKeySi&(A() As smeKey): On Error Resume Next: smeKeySi = UBound(A) + 1: End Function
Function smeKeyUB&(A() As smeKey): smeKeyUB = smeKeySi(A) - 1: End Function
Function smeTblDes(Lno, T As eSmeTblDes) As smeTblDes
With smeTblDes
    .Lno = Lno
    .T = T
End With
End Function
Function AddsmeTblDes(A As smeTblDes, B As smeTblDes) As smeTblDes(): PushsmeTblDes AddsmeTblDes, A: PushsmeTblDes AddsmeTblDes, B: End Function
Sub PushsmeTblDesAy(O() As smeTblDes, A() As smeTblDes): Dim J&: For J = 0 To smeTblDesUB(A): PushsmeTblDes O, A(J): Next: End Sub
Sub PushsmeTblDes(O() As smeTblDes, M As smeTblDes): Dim N&: N = smeTblDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smeTblDesSi&(A() As smeTblDes): On Error Resume Next: smeTblDesSi = UBound(A) + 1: End Function
Function smeTblDesUB&(A() As smeTblDes): smeTblDesUB = smeTblDesSi(A) - 1: End Function
Function smeTFDes(Lno, T As eSmeTFDes) As smeTFDes
With smeTFDes
    .Lno = Lno
    .T = T
End With
End Function
Function AddsmeTFDes(A As smeTFDes, B As smeTFDes) As smeTFDes(): PushsmeTFDes AddsmeTFDes, A: PushsmeTFDes AddsmeTFDes, B: End Function
Sub PushsmeTFDesAy(O() As smeTFDes, A() As smeTFDes): Dim J&: For J = 0 To smeTFDesUB(A): PushsmeTFDes O, A(J): Next: End Sub
Sub PushsmeTFDes(O() As smeTFDes, M As smeTFDes): Dim N&: N = smeTFDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smeTFDesSi&(A() As smeTFDes): On Error Resume Next: smeTFDesSi = UBound(A) + 1: End Function
Function smeTFDesUB&(A() As smeTFDes): smeTFDesUB = smeTFDesSi(A) - 1: End Function
Function smeFldDes(Lno, T As eSmeFldDes) As smeFldDes
With smeFldDes
    .Lno = Lno
    .T = T
End With
End Function
Function AddsmeFldDes(A As smeFldDes, B As smeFldDes) As smeFldDes(): PushsmeFldDes AddsmeFldDes, A: PushsmeFldDes AddsmeFldDes, B: End Function
Sub PushsmeFldDesAy(O() As smeFldDes, A() As smeFldDes): Dim J&: For J = 0 To smeFldDesUB(A): PushsmeFldDes O, A(J): Next: End Sub
Sub PushsmeFldDes(O() As smeFldDes, M As smeFldDes): Dim N&: N = smeFldDesUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function smeFldDesSi&(A() As smeFldDes): On Error Resume Next: smeFldDesSi = UBound(A) + 1: End Function
Function smeFldDesUB&(A() As smeFldDes): smeFldDesUB = smeFldDesSi(A) - 1: End Function
Function SchmPsr(Dta As SchmDta, Er As SchmEr) As SchmPsr
With SchmPsr
    .Dta = Dta
    .Er = Er
End With
End Function
