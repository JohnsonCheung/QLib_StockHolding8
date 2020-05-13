Attribute VB_Name = "MxIdeMdMdyLnDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde"
Const CMod$ = CLib & "MxIdeMdMdyLnDrs."
Const CNs$ = "MdyLn"
Public Const MdyLnFF$ = "OpLno LinOp OldL NewL"

Function MdyLnDy(Cur As LLn, Ept As LLn) As Variant()
Dim IsIns As Boolean, IsRpl As Boolean, IsDlt As Boolean
Dim InsLno%, RplLno%, DltMdln%
Dim NewL$, OldL$
    'Ins ---------------
    If Ept.Lno <> 0 Then
        If Cur.Lno <> Ept.Lno Then
            If Cur.Ln <> Ept.Ln Then
                IsIns = True
                InsLno = Ept.Lno
                NewL = Ept.Ln
            End If
        End If
    End If
    
    'Dlt ---------------
    If Ept.Lno = 0 Then
        If Cur.Lno <> 0 Then
            IsDlt = True
            DltMdln = Cur.Lno
            OldL = Cur.Ln
        End If
    End If
    
    'Rpl ---------------
    If Ept.Lno <> 0 Then
        If Ept.Lno = Cur.Lno Then
            If Ept.Ln <> Cur.Ln Then
                IsRpl = True
                RplLno = Ept.Lno
                NewL = Ept.Ln
                OldL = Cur.Ln
            End If
        End If
    End If
If IsIns Then PushI MdyLnDy, Array(InsLno, "Ins", "", NewL)
If IsDlt Then PushI MdyLnDy, Array(DltMdln, "Dlt", OldL, "")
If IsRpl Then PushI MdyLnDy, Array(RplLno, "Rpl", OldL, NewL)
End Function
