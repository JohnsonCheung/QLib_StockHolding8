Attribute VB_Name = "MxDaoAdoCatt"
Option Explicit
Option Compare Text
Type Catt: C As Catalog: T As Adox.Table: End Type
Function Catt(C As Catalog, T As Adox.Table) As Catt
With Catt
    Set .C = C
    Set .T = T
End With
End Function
Function CattzFxw(Fx$, Optional W$) As Catt
Dim C As Catalog, T As Adox.Table
Set C = CatzFx(Fx)
Set T = C.Tables(AxTbn(W))
CattzFxw = Catt(C, T)
End Function
Function CattzFbt(Fb$, T$) As Catt
Dim C As Catalog, Td As Adox.Table
Set C = CatzFb(Fb)
Set Td = C.Tables(T)
CattzFbt = Catt(C, Td)
End Function

