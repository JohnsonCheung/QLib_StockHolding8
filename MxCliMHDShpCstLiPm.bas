Attribute VB_Name = "MxCliMHDShpCstLiPm"
Option Compare Text
Option Explicit
Const CLib$ = "QRelCst."
Const CMod$ = CLib & "MxCliMHDShpCstLiPm."
Private Db As Database

Property Get ShpCstLnkPmLy() As String()
Const LnkColVblYYHT1$ = _
    " ZHT1   D Brand  |" & _
    " RateSc M Amount |" & _
    " VdtFm  M [Valid From]  |" & _
    " VdtTo  M [Valid to]"

Const LnkColVblzUom$ = _
    "Sku    M Material |" & _
    "Des    M [Material Description] |" & _
    "Sc_U   M SC |" & _
    "StkUom M [Base Unit of Measure] |" & _
    "Topaz  M [Topaz Code] |" & _
    "ProdH  M [Product hierarchy]"
 
Const LnkColVblzMB52$ = _
    " Sku    M Material |" & _
    " Whs    M Plant    |" & _
    " QInsp  D [In Quality Insp#]|" & _
    " QUnRes D Unrestricted|" & _
    " QBlk   D Blocked"
'A = "MB52": PushObj O, LiFDtoInfLnkColVbl(A, A, "Sheet1", LnkColVblzMB52)
'A = "UOM":  PushObj O, LiFDtoInfLnkColVbl(A, A, "Sheet1", LnkColVblzUom)
'            PushObj O, LiFDtoInfLnkColVbl("ZHT1", "ZHT18701", "8701", LnkColVblYYHT1)
'            PushObj O, LiFDtoInfLnkColVbl("ZHT1", "ZHT18601", "8601", LnkColVblYYHT1)
End Property
