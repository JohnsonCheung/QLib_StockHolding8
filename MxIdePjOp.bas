Attribute VB_Name = "MxIdePjOp"
Option Explicit
Option Compare Text
Const CNs$ = "Pj.Op"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdePjOp."

Sub BrwPjPthP()
BrwPth PthP
End Sub

Sub RmvPj(P As VBProject)
Const CSub$ = CMod & "RmvPj"
On Error GoTo X
Dim Pjn$: Pjn = P.Name
P.Collection.Remove P
Exit Sub
X:
Dim E$: E = Err.Description
WarnLn CSub, FmtQQ("Cannot remove P[?] Er[?]", Pjn, E)
End Sub

Private Sub SavPj__Tst()
Dim P As VBProject
Dim Wb As Workbook: Set Wb = NwWb
Dim Fx$: Fx = TmpFx
Wb.SaveAs Fx
Set P = PjzWb(VisWb(Wb))
Dim Cmp As VBComponent: Set Cmp = P.VBComponents.Add(vbext_ct_ClassModule)
Cmp.CodeModule.AddFromString "Sub AA()" & vbCrLf & "End Sub"
SavPj P
End Sub
Sub SavP(): SavPj CPj: End Sub
Sub SavPj(P As VBProject)
Const CSub$ = CMod & "SavPj"
If P.Saved Then
    Debug.Print FmtQQ("SavPj: Pj(?) is already saved", P.Name)
    Exit Sub
End If
'Chk Vbe
    Dim Vbe As Vbe
    Set Vbe = P.Collection.Vbe
    If ObjPtr(Vbe.ActiveVBProject) <> ObjPtr(P) Then Stop: Exit Sub
Dim Fnn$
    Fnn = PjFnn(P)
    If Fnn = "" Then
        Thw CSub, "Pj file name is blank.  The pj needs to saved first in order to have a pj file name", "Pj", P.Name
    End If
ActPj P
Dim B As CommandBarButton: Set B = IdeBtnSavzV(Vbe) 'Set Sav-Button to B
    If Not HasPfx(B.Caption, "&Save " & Fnn) Then Thw CSub, "Caption is not expected", "Save-Bottun-Caption Expected", B.Caption, "&Save " & Fnn
B.Execute '<===== Save
DoEvents
If Not P.Saved Then
    Debug.Print FmtQQ("SavPj: Pj(?) cannot be saved for unknown reason <=================================", P.Name)
Else
    Debug.Print FmtQQ("SavPj: Pj(?) is saved <---------------", P.Name)
End If
End Sub
