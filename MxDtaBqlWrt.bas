Attribute VB_Name = "MxDtaBqlWrt"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaBqlWrt."
':Fbka: :Ft #Fil-name-of-BacK-Apostrophe#
' ! It is a fil of ext *.bka.txt.  There 1-Sgn-Ln, 0-to-N-Rmk-Lines and 0-To-N-Tbl-Lines.
' ! Sgn-Ln is  **BackApostropheSeparatedFile**<Dsn>**, where <Dsn> is a dta-set-nm.
' ! Rmk-Lines are lines between Sgn-Ln and (fst-T-Ln or eof)
' ! 1-Tbl-Lines is  1-T-Ln, 0-to-N-TblRmk-Lines, 1-Fld-Ln and 0-to-N-Dta-Lines.
' ! T-Ln           is
' ! Rmk-Lines are lines before fst *-Ln.  Rmk are for all tbl in the :Fbka:  Each individual tbl does not have it own rmk
' ! Lines before fst *-Ln are Rmk.  Each gp of one-*-Ln & N-`-Ln is one tbl.
' ! *-Ln is a Ln wi fst chr is *, :Starl: #Star-Line#.  `-Ln is a Ln wi fst chr is `, :Bkal:, #BacK-Apostrophe-Ln#.
' ! The *-Ln is *<T>
' ! The fst `-Ln is :Scff
' ! The rst `-Ln is :dta
':T:   :S  #Table-Name#
':Scff: :FF #ShtTyChr-Colon-FF#  ! It is spc sep of :Scfld:.  It desc ty and fldn of the tbl.
'!It has first line as ShtTyscfQBLin.
'!It rest of lines are records."

Sub InsRszBql(R As DAO.Recordset, Bql$)
R.AddNew
Dim Ay$(): Ay = Split(Bql, "`")
Dim F As DAO.Field, J%
For Each F In R.Fields
    If Ay(J) <> "" Then
        F.Value = Ay(J)
    End If
    J = J + 1
Next
R.Update
End Sub
Function BqlzRs$(A As DAO.Recordset)
Dim O$(), F As DAO.Field
For Each F In A.Fields
    If IsNull(F.Value) Then
        PushI O, ""
    Else
        PushI O, Replace(Replace(F.Value, vbCr, ""), vbLf, " ")
    End If
Next
Dim L$: L = Jn(O, "`")
If L = "401`HD0V4FOF00C9ZT" Then Stop
BqlzRs = L

End Function

Private Sub WrtFbqlzDb__Tst()
Dim P$: P = TmpPthi
WrtFbqlzDb P, DutyDtaDb
BrwPth P
Stop
End Sub

Private Sub WrtFbqlzT__Tst()
Dim T$: T = TmpFt
WrtFbql T, DutyDtaDb, "PermitD"
BrwFt T
End Sub

Sub WrtFbqlzDb(Pth, D As Database)
WrtFbqlzTny Pth, D, Tny(D)
End Sub

Sub WrtFbqlzTny(Pth, D As Database, Tny$())
Dim T, P$
P = EnsPthSfx(Pth)
For Each T In Tny
    WrtFbql P & T & ".bql.txt", D
Next
End Sub

Sub WrtFbql(Fbql$, D As Database, Optional T0$)
Dim T$
    T = T0
    If T = "" Then T = TzFbql(Fbql)
Dim F%: F = FnoO(Fbql)
Dim R As DAO.Recordset
Set R = RszT(D, T)
Dim L$: L = ShtTyBqlzT(D, T)
Print #F, L
With R
    While Not .EOF
        Print #F, BqlzRs(R)
        .MoveNext
    Wend
    .Close
End With
Close #F
End Sub
