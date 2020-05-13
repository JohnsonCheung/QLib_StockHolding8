Attribute VB_Name = "gzIONam"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzIONam."

Function GitIPth$():              GitIPth = MB52IPthPm: End Function
Function GitIFx$(A As Ymd):        GitIFx = GitIPth & GitIFxFn(A):         End Function
Function GitIFxFn$(A As Ymd):    GitIFxFn = "Git " & YYmdStr(A) & ".xlsx": End Function
Function GitIFxMsg$(A As Ymd):  GitIFxMsg = IIf(Dir(GitIFx(A)) = "", "<==Missing", ""): End Function

Function MB52OPth$():     MB52OPth = AppOPth:     End Function
Function MB52LasIFx$(): MB52LasIFx = MB52IFx(LasOHYmd): End Function
Function MB52LasOFx$(): MB52LasOFx = MB52OFx(LasOHYmd): End Function

Function MB52IFx$(A As Ymd):       MB52IFx = MB52IPthPm & MB52IFxFn(A):      End Function
Function MB52IFxFn$(A As Ymd):   MB52IFxFn = "MB52 " & YYmdStr(A) & ".xlsx": End Function

Function MB52OFx$(A As Ymd):     MB52OFx = MB52OPth & MB52OFxFn(A):       End Function
Function MB52OFx1$(A As Ymd):   MB52OFx1 = CpyToPth1Pm & MB52OFxFn(A):    End Function
Function MB52OFx2$(A As Ymd):   MB52OFx2 = CpyToPth2Pm & MB52OFxFn(A):    End Function
Function MB52OFxFn$(A As Ymd): MB52OFxFn = "On Hand (MB52) " & YYmdStr(A) & ".xlsx": End Function
Function MB52Tp$():               MB52Tp = TpPthP & "On Hand Template.xlsx": End Function

Function ShOPth$():                           ShOPth = MB52OPth: End Function
Function ShTpFn$():                           ShTpFn = "Stock Holding Template.xlsx": End Function
Function ShTp$():                               ShTp = TpPthP & ShTpFn: End Function
Function ShOFxFn$(Co As Byte, A As Ymd):     ShOFxFn = FmtQQ("Stock Holding ?(?00).xlsx", YYmdStr(A), Co): End Function
Function ShOFxzCoYmd$(A As CoYmd):       ShOFxzCoYmd = ShOFx(A.Co, A.Ymd): End Function
Function ShOFx$(Co As Byte, A As Ymd):         ShOFx = ShOPth & ShOFxFn(Co, A): End Function
Function ShOFx1$(Co As Byte, A As Ymd):       ShOFx1 = CpyToPth1Pm & ShOFxFn(Co, A): End Function
Function ShOFx2$(Co As Byte, A As Ymd):       ShOFx2 = CpyToPth2Pm & ShOFxFn(Co, A): End Function

Function ZHT0IPth$(): ZHT0IPth = AppIPth: End Function
Function ZHT0IFx$():   ZHT0IFx = CPmv("ZHT0_InpFx"): End Function
Function ZHT0WFx$():   ZHT0WFx = Pth(CPmv("ZHT0_InpFx")) & "ZHT0 (Wrk).xlsx": End Function

Function LasMB52OFx$(): LasMB52OFx = MB52OFx(LasOHYmd): End Function
Sub OpnLasMB52OFx(): OpnFxMax LasMB52OFx: End Sub
