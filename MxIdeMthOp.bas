Attribute VB_Name = "MxIdeMthOp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthOp."
Sub CpyMthAs(M As CodeModule, Mthn, AsMthn)
Const CSub$ = CMod & "CpyMthAs"
If HasMthzM(M, AsMthn) Then Inf CSub, "AsMth exist.", "Mdn FmMth AsMth", Mdn(M), Mthn, AsMthn: Exit Sub
End Sub

Private Sub DltMth__Tst()
Const CSub$ = CMod & "Z_DltMth"
Dim Md As CodeModule
Const Mthn$ = "YYRmv1"
Dim Bef$(), Aft$()
Crt:
        Set Md = TmpMod
        AppLines Md, LineszVbl("|'sdklfsdf||'dsklfj|Property Get YYRmv1()||End Property||Function YYRmv2()|End Function||'|Sub SetYYRmv1(V)|End Property")
Tst:
        Bef = Src(Md)
        DltMth Md, Mthn
        Aft = Src(Md)

Insp:   Insp CSub, "DltMth Test", "Bef DltMth Aft", Bef, Mthn, Aft
Rmv:    RmvMd Md
End Sub

Sub MovMth(Mthn, ToMdn)
MovMthzM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthzM(Md As CodeModule, Mthn, ToMd As CodeModule)
CpyMth Mthn, Md, ToMd
DltMth Md, Mthn
End Sub

Function CdzEmpFun$(FunNm)
CdzEmpFun = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function CdzEmpSub$(Subn)
CdzEmpSub = FmtQQ("Sub ?()|End Sub", Subn)
End Function

Sub AddSub(Subn)
AppLines CMd, CdzEmpSub(Subn)
JmpMth Subn
End Sub

Sub AddFun(FunNm)
AppLines CMd, CdzEmpFun(FunNm)
JmpMth FunNm
End Sub

Sub CpyMth(Mthn, FmM As CodeModule, ToM As CodeModule)
If HasMthzM(ToM, Mthn) Then Thw CSub, "ToM has mthn FmM", "Mthn FmM ToM", Mthn, Mdn(FmM), Mdn(ToM)
ToM.AddFromString MthlzMN(FmM, Mthn)
End Sub

Sub CpyMthAsVer(M As CodeModule, Mthn, Ver%)
Const CSub$ = CMod & "CpyMthAsVer"
'Ret True if copied
Dim VerMthn$, Newl$, L$, Oldl$
If Not HasMthzM(M, Mthn) Then Inf CSub, "No from-mthn", "Md Mthn", Mdn(M), Mthn: Exit Sub
VerMthn = Mthn & "_Ver" & Ver
'NewL
    L = MthlzMN(M, Mthn)
    Newl = Replace(L, "Sub " & Mthn, "Sub " & VerMthn, Count:=1)
'Rpl
    RplMth M, VerMthn, Newl
End Sub
