Attribute VB_Name = "MxIdePjCmpOp"
Option Compare Text
Option Explicit
Const CNs$ = "Cmp.Op"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdePjCmpOp."

Sub AddClsnn(Clsnn$) 'To CPj
AddCmpnnzP Clsnn, vbext_ct_ClassModule, CPj
JmpCmpn T1(Clsnn)
End Sub

Sub AddCmpSfxzP(Sfx, P As VBProject)
If P.Protection = vbext_pp_locked Then Exit Sub
Dim C As VBComponent
For Each C In P.VBComponents
    RenCmp C, C.Name & Sfx
Next
End Sub

Sub AddCmpSfxP(Sfx)
AddCmpSfxzP Sfx, CPj
End Sub

Sub AddCmpzEmp(P As VBProject, Ty As vbext_ComponentType, Nm)
Const CSub$ = CMod & "AddCmpzEmp"
If HasCmpzP(P, Nm) Then InfLn CSub, "Cmpn exist in Pj", "Cmpn Pjn", Nm, P.Name: Exit Sub
P.VBComponents.Add(Ty).Name = Nm ' no CStr will break
End Sub

Sub AddCmpnnzP(Cmpnn$, T As vbext_ComponentType, P As VBProject)
Dim N: For Each N In ItrzSS(Cmpnn)
    AddCmpzEmp P, T, N
Next
End Sub

Sub AddCmpzL(P As VBProject, Cmpn, Srcl$)
AddCmpzEmp P, vbext_ct_StdModule, Cmpn
AppLines MdzP(P, Cmpn), Srcl
End Sub

Sub AddMd(Modnn$)
AddCmpnnzP Modnn, vbext_ct_StdModule, CPj
JmpCmpn T1(Modnn)
End Sub

Sub AddModnzP(P As VBProject, Modn)
AddCmpzEmp P, vbext_ct_StdModule, Modn
End Sub

Sub AppLines(M As CodeModule, Lines$)
If Lines = "" Then Exit Sub
M.InsertLines M.CountOfLines + 1, Lines '<=====
End Sub

Sub AppLineszoInf(M As CodeModule, Lines$)
Const CSub$ = CMod & "AppLineszoInf"
Dim Bef&, Aft&, Exp&, Cnt&
Bef = M.CountOfLines
AppLines M, Lines
Aft = M.CountOfLines
Cnt = LnCnt(Lines)
Exp = Bef + Cnt
If Exp <> Aft Then
    Thw CSub, "After copy line count are inconsistents, where [Md], [LnCnt-Bef-Cpy], [LnCnt-of-lines], [Exp-LnCnt-Aft-Cpy], [Act-LnCnt-Aft-Cpy], [Lines]", _
        Mdn(M), Bef, Cnt, Exp, Aft, Lines
End If
End Sub

Sub AppLy(M As CodeModule, Ly$())
AppLines M, JnCrLf(Ly)
End Sub

Sub ClrTmpMod()
Dim N
For Each N In TmpModNyzP(CPj)
    If HasPfx(Md(N), "TmpMod") Then RmvCmpzN N
Next
End Sub

Function DftMd(M As CodeModule) As CodeModule
If IsNothing(M) Then
   Set DftMd = CMd
Else
   Set DftMd = M
End If
End Function

Sub DltCmpzPjn(P As VBProject, Mdn)
If Not HasCmpzP(P, Mdn) Then Exit Sub
P.VBComponents.Remove P.VBComponents(Mdn)
End Sub

Sub EnsCls(P As VBProject, Clsn)
EnsCmp P, vbext_ct_ClassModule, Clsn
End Sub

Sub EnsCmp(P As VBProject, Ty As vbext_ComponentType, Nm)
If Not HasCmpzP(P, Nm) Then AddCmpzEmp P, Ty, Nm
End Sub

Sub EnsMod(P As VBProject, Modn)
EnsCmp P, vbext_ct_StdModule, Modn
End Sub

Sub EnsMdl(M As CodeModule, Mdl$)
Const CSub$ = CMod & "EnsModLines"
If Mdl = Srcl(M) Then Inf CSub, "Same module lines, no need to replace", "Mdn", Mdn(M): Exit Sub
RplMd M, Mdl
End Sub

Sub EnsMd(P As VBProject, Mdn)
EnsCmp P, vbext_ct_StdModule, Mdn
End Sub

Function HasCmpzN(Cmpn) As Boolean
HasCmpzN = HasCmpzP(CPj, Cmpn)
End Function

Sub RenCmp(A As VBComponent, NewNm$)
Const CSub$ = CMod & "RenCmp"
If HasCmpzN(NewNm) Then
    InfLn CSub, "New cmp exists", "OldCmp NewCmp", A.Name, NewNm
Else
    A.Name = NewNm
End If
End Sub

Sub RmvCmp(A As VBComponent): A.Collection.Remove A: End Sub
Sub RmvCmpzN(Cmpn): RmvCmp Cmp(Cmpn): End Sub
Sub RmvMd(MdDn): RmvMdzM Md(MdDn): End Sub
Sub RmvMdzM(M As CodeModule)
'Debug.Print FmtQQ("RmvMd: Before Md(?) is deleted from Pj(?)", M, P)
M.Parent.Collection.Remove M.Parent
If True Then
    Dim N$, P$
    N = Mdn(M)
    P = PjnzM(M)
    Debug.Print FmtQQ("RmvMdzM: Md(?) is deleted from Pj(?)", N, P)
End If
End Sub

Sub RmvMdByPfx(P As VBProject, Pfx$)
Const CSub$ = CMod & "RmvMdByPfx"
Dim Ny$(): Ny = AwPfx(Mdny(P), Pfx)
If Si(Ny) = 0 Then InfLn CSub, FmtQQ("no module in Pj[?] begins with pfx-" & P.Name, Pfx): Exit Sub
Brw Ny, "RmvMdByPFx_"
If Cfm("Rmv those Md as show in the notepad?") Then
    Dim N: For Each N In Ny
        RmvMd Md(N)
    Next
End If
End Sub
Sub RplMdzPjSrc(P As VBProject, S As PjSrc) '#Rpl-Md-By-Dic#
Dim J%: For Each J In MdSrcUB(S.Md)
    RplMd P.VBComponents(Mdn).CodeModule, S.Md(J).Srcl
Next
End Sub

Private Sub RplMd__Tst()
Dim M As CodeModule: Set M = Md("QDao_Def_NewTd")
RplMd M, Srcl(M) & vbCrLf & "'"
End Sub
Function RplMd(M As CodeModule, Newl$) As Boolean
Brw Newl: Exit Function
Dim Mdn$: Mdn = M.Name
Dim Oldl$: Oldl = Srcl(M)
Dim IsSam As Boolean: IsSam = RTrimLines(Oldl) = RTrimLines(Newl)
W1ShwMsg IsSam, Oldl, Newl, Mdn
If IsSam Then Exit Function
ClrMd M
M.InsertLines 1, Newl
RplMd = True
End Function
Private Sub W1ShwMsg(IsSam As Boolean, Oldl$, Newl$, Mdn$)
Dim Msg$
    Dim OldC As String * 4: RSet OldC = LnCnt(Oldl)
    Msg = Replace("RplMd: OldCnt(?) ", "?", OldC)
    If IsSam Then
        Msg = Msg & "             " & Mdn & vbTab & "<--- Same"
    Else
        Dim NewC As String * 4: RSet NewC = LnCnt(Newl)
        Msg = Msg & Replace("NewCnt(?) ", "?", NewC) & Mdn
    End If
    Debug.Print Msg
End Sub

Sub RenModPfx(FmPfx$, ToPfx$): RenModPfxzP CPj, FmPfx, ToPfx: End Sub
Sub RenModPfxzP(Pj As VBProject, FmPfx$, ToPfx$)
Dim C As VBComponent, N$
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If HasPfx(C.Name, FmPfx) Then
            RenCmp C, RplPfx(C.Name, FmPfx, ToPfx)
        End If
    End If
Next
End Sub

Function SetCmpNm(A As VBComponent, Nm, Optional Fun$ = "SetCmpNm") As VBComponent
Dim Pj As VBProject
Set Pj = PjzCmp(A)
If HasCmpzP(Pj, Nm) Then
    Thw Fun, "Cmp already Has", "Cmp Has-in-Pj", Nm, Pj.Name
End If
If Pj.Name = Nm Then
    Thw Fun, "Cmpn same as Pjn", "Cmpn", Nm
End If
A.Name = Nm
Set SetCmpNm = A
End Function

Sub ChgToCls(FmModn$)
Const CSub$ = CMod & "ChgToCls"
If Not HasCmp(FmModn) Then InfLn CSub, "Mod not exist", "Mod", FmModn: Exit Sub
If Not IsMod(Md(FmModn)) Then InfLn CSub, "It not Mod", "Mod", FmModn: Exit Sub
Dim T$: T = Left(FmModn & "_" & Format(Now, "HHMMDD"), 31)
Md(FmModn).Name = T
AddClsnn FmModn
Md(FmModn).AddFromString Srcl(Md(T))
RmvCmpzN T
End Sub
