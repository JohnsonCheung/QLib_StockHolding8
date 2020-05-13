VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Switchboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_Switchboard."
Public mBool As Boolean
Private Sub Form_Open(Cancel As Integer)
SetOption "Confirm Action Queries", False
SetOption "Confirm Record Changes", False

' Minimize the database window and initialize the form.

    ' Move to the switchboard page that is marked as the default.
    Me.Filter = "[ItemNumber] = 0 AND [Argument] = 'Default' "
    Me.FilterOn = True
    DoCmd.Maximize
End Sub

Private Sub Form_Current()
' Update the caption and fill in the list of options.

    Me.Caption = Nz(Me![ItemText], "")
    FillOptions
    
End Sub

Private Sub FillOptions()
' Fill in the options for this switchboard page.

    ' The number of buttons on the form.
    Const conNumButtons = 8
    
    Dim dbs As Database
    Dim Rst As Recordset
    Dim strSQL As String
    Dim intOption As Integer
    
    ' Set the focus to the first button on the form,
    ' and then hide all of the buttons on the form
    ' but the first.  You can't hide the field with the focus.
    Me![Option1].SetFocus
    For intOption = 2 To conNumButtons
        Me("Option" & intOption).Visible = False
        Me("OptionLabel" & intOption).Visible = False
    Next intOption
    
    ' Open the table of Switchboard Items, and find
    ' the first item for this Switchboard Page.
    Set dbs = CurrentDb()
    strSQL = "SELECT * FROM [Switchboard Items]"
    strSQL = strSQL & " WHERE [ItemNumber] > 0 AND [SwitchboardID]=" & Me![SwitchboardID]
    strSQL = strSQL & " ORDER BY [ItemNumber];"
    Set Rst = dbs.OpenRecordset(strSQL)
    
    ' If there are no options for this Switchboard Page,
    ' display a message.  Otherwise, fill the page with the items.
    If (Rst.EOF) Then
        Me![OptionLabel1].Caption = "There are no items for this switchboard page"
    Else
        While (Not (Rst.EOF))
            Me("Option" & Rst![ItemNumber]).Visible = True
            Me("OptionLabel" & Rst![ItemNumber]).Visible = True
            Me("OptionLabel" & Rst![ItemNumber]).Caption = Rst![ItemText]
            Rst.MoveNext
        Wend
    End If

    ' Close the recordset and the database.
    Rst.Close
    dbs.Close

End Sub

Private Function HandleButtonClick(intBtn As Integer)
' This function is called when a button is clicked.
' intBtn indicates which button was clicked.

    ' Constants for the commands that can be executed.
    Const conCmdGotoSwitchboard = 1
    Const conCmdOpenFormAdd = 2
    Const conCmdOpenFormBrowse = 3
    Const conCmdOpenReport = 4
    Const conCmdCustomizeSwitchboard = 5
    Const conCmdExitApplication = 6
    Const conCmdRunMacro = 7
    Const conCmdRunCode = 8

    ' An error that is special cased.
    Const conErrDoCmdCancelled = 2501
    
    Dim dbs As Database
    Dim Rst As Recordset

On Error GoTo HandleButtonClick_Err

    ' Find the item in the Switchboard Items table
    ' that corresponds to the button that was clicked.
    Set dbs = CurrentDb()
    Set Rst = dbs.OpenRecordset("Switchboard Items", dbOpenDynaset)
    Rst.FindFirst "[SwitchboardID]=" & Me![SwitchboardID] & " AND [ItemNumber]=" & intBtn
    
    ' If no item matches, report the error and exit the function.
    If (Rst.NoMatch) Then
        MsgBox "There was an error reading the Switchboard Items table."
        Rst.Close
        dbs.Close
        Exit Function
    End If
    
    Select Case Rst![Command]
        
        ' Go to another switchboard.
        Case conCmdGotoSwitchboard
            Me.Filter = "[ItemNumber] = 0 AND [SwitchboardID]=" & Rst![Argument]
            
        ' Open a form in Add mode.
        Case conCmdOpenFormAdd
            DoCmd.OpenForm Rst![Argument], , , , acAdd

        ' Open a form.
        Case conCmdOpenFormBrowse
            DoCmd.OpenForm Rst![Argument]

        ' Open a report.
        Case conCmdOpenReport
            DoCmd.OpenReport Rst![Argument], acPreview

        ' Customize the Switchboard.
        Case conCmdCustomizeSwitchboard
            ' Handle the case where the Switchboard Manager
            ' is not installed (e.g. Minimal Install).
            On Error Resume Next
            Application.Run "ZZMAIN80.sbm_Entry"
            If (Err <> 0) Then MsgBox "Command not available."
            On Error GoTo 0
            ' Update the form.
            Me.Filter = "[ItemNumber] = 0 AND [Argument] = 'Default' "
            Me.Caption = Nz(Me![ItemText], "")
            FillOptions

        ' Exit the application.
        Case conCmdExitApplication
            CloseCurrentDatabase

        ' Run a macro.
        Case conCmdRunMacro
            DoCmd.RunMacro Rst![Argument]

        ' Run code.
        Case conCmdRunCode
            Application.Run Rst![Argument]

        ' Any other command is unrecognized.
        Case Else
            MsgBox "Unknown option."
    
    End Select

    ' Close the recordset and the database.
    Rst.Close
    dbs.Close
    
HandleButtonClick_Exit:
    Exit Function

HandleButtonClick_Err:
    ' If the action was cancelled by the user for
    ' some reason, don't display an error message.
    ' Instead, resume on the next line.
    If (Err = conErrDoCmdCancelled) Then
        Resume Next
    Else
        MsgBox "There was an error executing the command.", vbCritical
        Resume HandleButtonClick_Exit
    End If
    
End Function
