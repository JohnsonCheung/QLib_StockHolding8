Attribute VB_Name = "MxVbRun"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "InterAct"
Const CMod$ = CLib & "MxVbRun."
Enum EmWaitRslt
    EiTimUp
    EiCnl
End Enum
Type WaitOpt
    TimOutSec As Integer
    ChkSec As Integer
    KeepFcmd As Boolean
End Type
Declare Function GetCurrentProcessId& Lib "Kernel32.dll" ()
'Declare Function GetProcessId& Lib "Kernel32.dll" (ProcessHandle&)
'Const Ps1Str$ = "function Get-ExcelProcessId { try { (Get-Process -Name Excel).Id } finally { @() } }" & vbCrLf & _
'"Stop-Process -Id (Get-ExcelProcessId)"

Function WaitOpt(TimOutSec%, ChkSec%, KeepFcmd As Boolean) As WaitOpt
With WaitOpt
.TimOutSec = TimOutSec
.ChkSec = ChkSec
.KeepFcmd = KeepFcmd
End With
End Function

Property Get DftWait() As WaitOpt
DftWait = WaitOpt(30, 5, False)
End Property

Sub KillProcessId(ProcessId&)
End Sub

Function RunFps1&(Fps1$, Optional PmStr$)
RunFps1 = RunFcmd("PowerShell", QuoDbl(Fps1) & " " & PmStr)
End Function

Function RunFcmd&(Fcmd$, Optional PmStr$, Optional Sty As VbAppWinStyle = vbMaximizedFocus)
Dim Ln
    Ln = QuoDbl(Fcmd) & AddPfxSpcIfNB(PmStr)
RunFcmd = Shell(Ln, Sty)
End Function

Function WaitFcmdw(Fcmdw$, W As WaitOpt, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus) As Boolean _
'Return True, if Fcmdw has generated the Fwaitg
Dim ProcessId&: ProcessId = Shell(Fcmdw, Sty)
Dim Fw$: Fw = Fwaitg(Fcmdw)
If WaitFwaitg(Fw, W.ChkSec, W.TimOutSec) Then
    Kill Fw
    WaitFcmdw = True
Else
    KillProcessId ProcessId
End If
If Not W.KeepFcmd Then Kill Fcmdw
End Function

Function WaitFwaitg(Fwaitg$, Optional ChkSec% = 10, Optional TimOutSec% = 60, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus) As Boolean _
'Return True, if Fwaitg is found.
Dim J%
For J = 1 To TimOutSec \ ChkSec
    If HasFfn(Fwaitg) Then
        Kill Fwaitg
        Exit Function
    End If
    If Not Wait(ChkSec%) Then Exit Function
Next
End Function

Private Sub Fcmdw__Tst()
Debug.Print LineszFt(Fcmdw("Dir"))
End Sub
Function Fwaitg$(Fcmd$)
Fwaitg = Fcmd & ".wait.txt"
End Function

Function Fcmdw$(CmdLines$)
Dim T$: T = TmpFcmd
Dim EchoLin: EchoLin = FmtQQ("Echo > ""?""", Fwaitg(T))
Dim S$: S = CmdLines & vbCrLf & EchoLin
Fcmdw = WrtStr(S, T)
End Function

Private Sub RunFcmd__Tst()
RunFcmd "Cmd"
MsgBox "AA"
End Sub

Function Wait(Optional Sec% = 1) As EmWaitRslt
Dim Till As Date: Till = AftSec(Sec)
Wait = IIf(Xls.Wait(Till), EiTimUp, EiCnl)
End Function

Function AftSec(Sec%) As Date 'Return the Date after Sec from Now
AftSec = DateAdd("S", Sec, Now)
End Function
Function Pipe(Pm, Mthnn$)
Dim O: Asg Pm, O
Dim I
For Each I In Ny(Mthnn)
   Asg Run(I, O), O
Next
Asg O, Pipe
End Function

Function RunAvzIgnEr(Mthn, Av())
Const CSub$ = CMod & "RunAvzIgnEr"
If Si(Av) > 9 Then Thw CSub, "Si(Av) should be 0-9", "Si(Av)", Si(Av)
On Error Resume Next
RunAv Mthn, Av
End Function

Function RunAv(Mthn, Av())
Const CSub$ = CMod & "RunAv"
Dim O
Select Case Si(Av)
Case 0: O = Run(Mthn)
Case 1: O = Run(Mthn, Av(0))
Case 2: O = Run(Mthn, Av(0), Av(1))
Case 3: O = Run(Mthn, Av(0), Av(1), Av(2))
Case 4: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case 9: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7), Av(8))
Case Else: Thw CSub, "UB-Av should be <= 8", "UB-Si Mthn", UB(Av), Mthn
End Select
RunAv = O
End Function

