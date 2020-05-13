Attribute VB_Name = "MxVbFsPthOpClrR"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsPthOpClrR."

Function IsErClrPthR(Pth) As Boolean
If NoPth(ParPth(Pth)) Then
    MsgBox "Following path not found" & vbCrLf & ParPth(Pth), vbCritical + vbDefaultButton1
    IsErClrPthR = True
    Exit Function
End If
If NoPth(Pth) Then Exit Function
DltPthR Pth
If HasPth(Pth) Then
    MsgBox "Cannot clear the following path.  May be due some file or sub-folders in the path is openned.  Close them and re-try." & vbCrLf & vbCrLf & Pth, vbDefaultButton1 + vbCritical
    IsErClrPthR = True
    Exit Function
End If
End Function

Function ResPthA$()
ResPthA = ResHom & "A"
End Function

Function ResPthB$()
ResPthB = ResHom & "B"
End Function

Sub CrtResPthA()
Dim P$: P = EnsResPseg("A\Lvl1-A\B\C\")
WrtRes "AA", P & "AA.Txt"
WrtRes "abc", P & "ABC.Txt"
WrtStr "AA", P & "AA.txt"
WrtStr "abc", P & "ABC.Txt"
End Sub

Private Sub ClrPthR__Tst()
CrtResPthA
Dim T$: T = ResPthA
BrwPth T
Stop
Debug.Print IsCfmAndClrPthR(T)
End Sub

Sub DltPthR(Pth)
Dim F: For Each F In Itr(FfnAyR(Pth))
    DltFfn F
Next
DltEmpPthR Pth
DltPthSilent Pth
End Sub

Function IsCfmAndClrPthR(Pth) As Boolean
If IsCfmClrPthR(Pth) Then
    IsCfmAndClrPthR = IsErClrPthR(Pth)
End If
End Function

Function IsCfmClrPthR(Pth) As Boolean
Const CSub$ = CMod & "IsCfmClrPthR"
If NoPth(ParPth(Pth)) Then Thw CSub, "Path not found", "Pth", Pth
If NoPth(Pth) Then IsCfmClrPthR = True: Exit Function
If MsgBox(Pth & vbCrLf & vbCrLf & "In next prompt, Input [Yes], to DELETE" & vbCrLf & "All files and folders under above path and the path itself.", vbDefaultButton1 + vbYesNo + vbQuestion) <> vbYes Then Exit Function
Dim A$: A = InputBox("Input [YES] to delete all files and folders under path in previous prompt." & vbCrLf & _
"After delete, CANNOT un-delete")
If A <> "YES" Then Exit Function
IsCfmClrPthR = True
End Function
