Attribute VB_Name = "MxIdeMdOpLnOp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeMdLnOp."

Sub AppCdl(M As CodeModule, Cdl$)
Debug.Print "Begin AppCdl....."
M.InsertLines M.CountOfLines + 1, Cdl
Debug.Print "End AppCdl......"
End Sub
Sub RplCdl(M As CodeModule, Lno&, Cdl$) ' Replace modul
Dim Ln$: Ln = M.Lines(Lno, 1)
M.ReplaceLine Lno, Cdl
Inf CSub, "A line is replaced by Mdl", "Mdn Lno [the line being replaced] [by Cdl]", Mdn(M), Lno, Ln, Cdl
End Sub

Sub DltCdl(M As CodeModule, Lno&, Optional Cnt = 1)
ChkMdLno CSub, M, Lno
If Cnt <= 0 Then PmEr CSub, "Given Cnt should be >=1", "Cnt", Cnt
W1Inf CSub, M, Lno, Cnt
M.DeleteLines Lno, Cnt
End Sub
Private Sub W1Inf(Fun$, M As CodeModule, Lno&, Cnt)
Dim Cdl$: Cdl = M.Lines(Lno, Cnt)
Inf Fun, "Cdl is deleted.", "Mdn Lno Cnt Cdl", Mdn(M), Lno, Cnt, Cdl
End Sub

Sub InsCdl(M As CodeModule, Lno&, Cdl$)
M.InsertLines Lno, Cdl
Debug.Print FmtQQ("InsCdl: Line is ins Lno[?] Md[?] Ln[?]", Lno, Mdn(M), Cdl)
End Sub

Sub DltCdlzLcnt(M As CodeModule, A As Lcnt)
DltCdl M, A.Lno, A.Cnt
End Sub
