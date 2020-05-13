Attribute VB_Name = "MxAcsFun"
Option Compare Text
Option Explicit
Const CLib$ = "QAcs."
Const CNs$ = "Acs.Fun"
Const CMod$ = CLib & "MxAcsFun."


Function HasFb(A As Access.Application, Fb) As Boolean
On Error GoTo X
HasFb = A.CurrentDb.Name = Fb: Exit Function
X:
End Function

Function GetAcs() As Access.Application
Set GetAcs = GetObject(, "Access.Application")
End Function

Function Acs() As Access.Application
Set Acs = Access.Application
End Function

Function AcszDb(Db As Database) As Access.Application
Set AcszDb = AcszFb(Db.Name)
End Function

Function IsOkAcs(A As Access.Application) As Boolean
On Error GoTo X
Dim N$: N = A.Name
IsOkAcs = True
Exit Function
X:
End Function

Function AcszFb(Fb, Optional IsExl As Boolean) As Access.Application
Dim O As Access.Application: Set O = NwAcs
O.OpenCurrentDatabase Fb, IsExl
Set AcszFb = O
End Function

Function DbnzAcs$(A As Access.Application)
On Error Resume Next
DbnzAcs = A.CurrentDb.Name
End Function

Function DftAcs(A As Access.Application) As Access.Application
'Ret :@A if Not Nothing or :NwAcs
If IsNothing(A) Then
    Set DftAcs = NwAcs
Else
    Set DftAcs = A
End If
End Function

Function FbzAcs$(A As Access.Application)
'Ret :Dbn openned in @A or *Blnk
On Error Resume Next
FbzAcs = A.CurrentDb.Name
End Function

Sub MinvAcs(A As Access.Application)
A.Visible = True
MiniAcs A
End Sub

Sub MiniAcs(A As Access.Application)
A.DoCmd.RunCommand acCmdAppMinimize
End Sub

Sub MaxiAcs(A As Access.Application)
A.DoCmd.RunCommand acCmdAppMaximize
End Sub

Function NwAcs() As Access.Application
Dim O As Access.Application: Set O = CreateObject("Access.Application")
O.Visible = True
MiniAcs O
Set NwAcs = O
End Function

Function PjzAcs(A As Access.Application) As VBProject
Set PjzAcs = A.Vbe.ActiveVBProject
End Function

Function PjzFba(Fba, A As Access.Application) As VBProject
OpnFb A, Fba
Set PjzFba = PjzAcs(A)
End Function

Sub QuitAcs(A As Access.Application)
If IsNothing(A) Then Exit Sub
On Error Resume Next
Stamp "QuitAcs: Begin"
Stamp "QuitAcs: Cls":         A.CloseCurrentDatabase
Stamp "QuitAcs: Quit":        A.Quit
Stamp "QuitAcs: Set Nothing": Set A = Nothing
Stamp "QuitAcs: End"
End Sub

Sub SavRec()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Function ShwAcs(A As Access.Application) As Access.Application
If Not A.Visible Then A.Visible = True
Set ShwAcs = A
End Function

Function CvAcs(A) As Access.Application
Set CvAcs = A
End Function
