Attribute VB_Name = "MxVbRunMapDrive"
Option Compare Text
Option Explicit
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxVbRunMapDrive."

Sub MapDrive(Drv$, Pth$)
RmvDrive Drv
Shell FmtQQ("Subst ? ""?""", Drv, Pth)
End Sub

Sub MapNDrive()
MapDrive "N:", "c:\users\user\desktop\Mhd"
End Sub

Sub RmvDrive(Drv$)
Shell "Subst /d " & Drv
End Sub

Sub RmvNDrive()
RmvDrive "N:"
End Sub
