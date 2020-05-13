Attribute VB_Name = "MxVbDDE"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "InterAct"
Const CMod$ = CLib & "MxVbDDE."

Sub TstDDE()
Dim A As New Excel.Application
'Dim A As Excel.Application: Set A = GetObject(, "Excel.Application")
A.Visible = True
A.Workbooks.Add
Dim Pj As VBProject: Set Pj = A.Vbe.VBProjects.Item(1)
Dim Cmp As VBComponent: Set Cmp = Pj.VBComponents.Add(vbext_ct_StdModule)
If True Then
    Dim ChannelNumber&: ChannelNumber = Application.DDEInitiate( _
        Application:="Excel", _
        topic:="Book1")
    VBA.AppActivate "Book1"
    Application.DDEExecute ChannelNumber, "%{F11}"
    Application.DDETerminate ChannelNumber
    MsgBox "AA"
Else
    A.SendKeys "%{F11}"
End If
End Sub
