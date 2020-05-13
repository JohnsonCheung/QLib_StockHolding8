Attribute VB_Name = "MxIdeSrcShf"
Option Explicit
Option Compare Text

Function ShfRmk$(OLin$)
Dim L$
L = LTrim(OLin)
If FstChr(L) = "'" Then
    ShfRmk = Mid(L, 2)
    OLin = ""
End If
End Function

Function ShfAs(OLn$) As Boolean
ShfAs = ShfTermX(OLn, "As")
End Function

Function ShfNmAftAs$(OLn$) ' Shf the name aft " As ", if no name aft as, thw er
If Not ShfAs(OLn) Then Exit Function
Dim O$: O = ShfNm(OLn)
If O = "" Then Thw CSub, "no name after as", OLn
ShfNmAftAs = O
End Function

Function ShfDclSfx$(OLin$)
Dim O$: O = ShfTyChr(OLin)
If O <> "" Then
    ShfDclSfx = O & IIf(ShfBkt(OLin), "()", "")
    Exit Function
End If
Dim Bkt$:
    If ShfBkt(OLin) Then
        Bkt = "()"
    End If
If ShfAs(OLin) Then
    Dim DNm$: DNm = ShfDotNm(OLin):
    ShfDclSfx = Bkt & " As " & DNm
    If DNm = "" Then Stop
Else
    ShfDclSfx = Bkt
End If
End Function

Function ShfDim(OLin$) As Boolean
ShfDim = ShfTerm(OLin) = "Dim"
End Function
Function ShfTermAftAs$(OLin$)
If Not ShfTermX(OLin, "As") Then Exit Function
ShfTermAftAs = ShfTerm(OLin)
End Function

Function ShfShtMdy$(OLin$)
ShfShtMdy = ShtMdy(ShfMdy(OLin))
End Function

Function ShfShtMthTy$(OLin$)
ShfShtMthTy = ShtMthTy(ShfMthTy(OLin))
End Function

Function ShfShtMthKd$(OLin$)
ShfShtMthKd = ShtMthKdzShtMthTy(ShtMthTy(ShfMthTy(OLin)))
End Function

Function ShfSub(OLin$) As Boolean
ShfSub = ShfPfx(OLin, "Sub ")
End Function

Function ShfPrv(OLin$) As Boolean
ShfPrv = ShfPfx(OLin, "Private ")
End Function

Function ShfMdy$(OLn$)
ShfMdy = Mdy(OLn)
OLn = LTrim(RmvPfx(OLn, ShfMdy))
End Function

Function ShfKd$(OLin$)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
ShfKd = T
OLin = LTrim(RmvPfx(OLin, T))
End Function

Function ShfMthSfx$(OLin$)
ShfMthSfx = ShfChr(OLin, TyChrLis)
End Function
