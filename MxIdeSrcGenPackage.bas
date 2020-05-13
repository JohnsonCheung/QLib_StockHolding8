Attribute VB_Name = "MxIdeSrcGenPackage"
Option Explicit
Option Compare Text
Const CLib$ = "QPackage."
Const CMod$ = CLib & "MxIdeSrcGenPackage."
Public Const Libnn$ = "QAcs QApp QClient QDao QDta QGit QIde QItrObj QShpCst QSql QSudoku QTp QVb QXls QZip"
Public Const QAcsMdnn$ = ""
Public Const QAppMdnn$ = ""
Public Const QDaoMdnn$ = ""
Public Const QDtaMdnn$ = ""
Public Const QGitMdnn$ = ""
Public Const QIdeMdnn$ = ""
Public Const QSqlMdnn$ = ""
Public Const QSudokuMdnn$ = ""
Public Const QTpMdnn$ = ""
Public Const QVbMdnn$ = ""
Public Const QXlsMdnn$ = ""
Public Const QZipMdnn$ = ""

Function CLibvAy() As String() '
CLibvAy = SyzSS(CLibvv)
End Function

Private Function MdnnDi() As Dictionary 'A di of CPj mapping CLib->Mdnn
Static X As Boolean, Y As New Dictionary
If Not X Then
    X = True
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
    Y.Add "QDta", QDtaMdnn
End If
Set MdnnDi = Y
End Function

Function CModvv$(Libn$)
CModvv = JnSpc(CModvAy(CPj))
End Function

Function CModvAy(Pj As VBProject) As String()
Dim C As VBComponent: For Each C In Pj.VBComponents
    PushI CModvAy, CModv(Dcl(C.CodeModule))
Next
End Function

Function CModNy$(Pj As VBProject, Libn$)
Dim O$()
Dim C As VBComponent: For Each C In Pj.VBComponents
    Dim D$(): D = Dcl(C.CodeModule)
    If HasCLibv(D, Libn) Then
        PushNBNDup O, CModv(Dcl(C.CodeModule))
    End If
Next
End Function

Function MdnnzLibn$(Libn$)
MdnnzLibn = MdnnDi(Libn)
End Function

Function CLibvv$()
CLibvv = SrtSS(CLibvvzP(CPj))
End Function
Function CLibvvzP$(P As VBProject)
Dim O$()
Dim C As VBComponent: For Each C In P.VBComponents
    PushNBNDup O, CLibv(Dcl(C.CodeModule))
Next
CLibvvzP = JnSpc(SrtAy(O))
End Function
