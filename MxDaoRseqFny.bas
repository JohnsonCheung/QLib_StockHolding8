Attribute VB_Name = "MxDaoRseqFny"
Option Explicit
Option Compare Text
Const CNs$ = "Fny.Rseq"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDaoRseqFny."

Function RseqFnyEnd(Fny$(), EndFny$()) As String()
Dim OkEnd$(): OkEnd = IntersectAy(Fny, EndFny)
RseqFnyEnd = AddSy(MinusSy(Fny, OkEnd), OkEnd)
End Function

Function RseqFnyFront(Fny$(), FrontFny$()) As String()
Dim Front$(): Front = IntersectAy(FrontFny, Fny)
RseqFnyFront = AddSy(Front, MinusSy(Fny, Front))
End Function
