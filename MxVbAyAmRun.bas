Attribute VB_Name = "MxVbAyAmRun"
Option Explicit
Option Compare Text

Function AmPX(IInto, Ay, PX$, P) '#Ay-Map-PX# ret an array by running Fun-@Px which takes 2 Pm-(@P *X) where *X is ele of @Ay
'@IInto:Cml :IAy #Item-Into# which is to create the return ay
Dim O: O = NwAy(IInto)
Dim I: For Each I In Itr(Ay)
    PushI O, Run(PX, P, I)
Next
End Function

Function AmXP(IInto, Ay, XP$, P) '#Ay-Map-XP# ret an array by running Fun-@XP which takes 2 Pm-(*X @P) where *X is ele of @Ay
'@IInto:Cml :IAy #Item-Into# which is to create the return ay
Dim O: O = NwAy(IInto)
Dim I: For Each I In Itr(Ay)
    PushI O, Run(XP, I, P)
Next
End Function
