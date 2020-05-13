Attribute VB_Name = "MxTpSpecEu"
Option Explicit
Option Compare Text
Type SpecEu
    Top() As String  ' no --- in front.  This will be put the top of TpLy after adding ---.
    LnEnd() As ILn    ' no --- in front of x.Ln.  This will be put the end of line of x.Ix after adding ---
End Type

Function FmtSpecEu(Ly$(), E As SpecEu) As String()
Stop
Dim Bdy$(): Bdy = Ly
    Dim J%: For J = 0 To ILnUB(E.LnEnd)
        Dim L As ILn: L = E.LnEnd(J)
        Bdy(L.Ix) = Bdy(L.Ix) & " --- " & L.Ln
    Next
FmtSpecEu = AddSy(E.Top, Bdy)
End Function

