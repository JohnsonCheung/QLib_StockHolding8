Attribute VB_Name = "MxVbGit"
Option Compare Text
Option Explicit
Const CLib$ = "QGit."
Const CMod$ = CLib & "MxVbGit."
Const Fgit$ = "C:\Program Files\Git\Cmd\git.Exe"
Const FgitQ$ = vbDblQ & Fgit & vbDblQ
Function Fcmd$(CmdPfx$, CmdCd$)
Fcmd = WrtStr(CmdCd, TmpFcmd(CmdPfx))
End Function
Function CmitCmdCd$(Optional Msg$ = "Commit", Optional ReInit As Boolean)
BfrClr
    Dim P$:     P = SrcPthP
    BfrV "Cd """ & P & """"
    If ReInit Then X "Rd .git /s/q"
    BfrV FmtQQ("? init", FgitQ)       'If already init, it will do nothing
    BfrV FmtQQ("? add -A", FgitQ)
    BfrV FmtQQ("? commit -m ""?""", FgitQ, Msg)
    BfrV "Pause"
CmitCmdCd = BfrLines
End Function
Sub Cmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
Dim CmdCd$: CmdCd = CmitCmdCd(Msg, ReInit)
ShellMax Fcmd("Cmit", CmdCd)
End Sub
Function PushGitCmdCd$()
BfrClr
    Dim P$: P = SrcPthP
    BfrV FmtQQ("Cd ""?""", P)
    BfrV FmtQQ("? push -u https://johnsoncheung@github.com/johnsoncheung/?.git master", FgitQ, PjnzSrcPth(P))
    BfrV "Pause ....."
PushGitCmdCd = BfrLines
End Function

Sub GitPush()
ShellMax Fcmd("Push", PushGitCmdCd)
End Sub

Function HasInternet() As Boolean
Stop
End Function

Function PjnzSrcPth$(SrcPth$)
Const CSub$ = CMod & "PjnzSrcPth"
Dim P$: P = RmvPthSfx(SrcPth)
If Ext(P) <> ".src" Then Thw CSub, "Not source path", "CmitgPth", SrcPth
PjnzSrcPth = RmvExt(Fn(P))
End Function
