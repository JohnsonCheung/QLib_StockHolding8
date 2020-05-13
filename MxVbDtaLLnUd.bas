Attribute VB_Name = "MxVbDtaLLnUd"
Option Explicit
Option Compare Text
#If Doc Then
'Udt Spec
'  @Spect snd Term of fst-line
'  @Specn  third Term of fst-line
'  @ShtRmk Rst aft third term
'  @Rmk    all indented lines after the fst-ln
'          any dash dash pfx will be removed
'  fst-ln  fst term must be *Spec, otherwise, error.
'Udt Speci Spec-Item
'  @Specit Fst Term of hdr-line
'  @Specin  Snd Term
'  @ShtRmk  Rst of aft snd term
'  @Rmk     All indent Ly
'  @LLn     following of hdr-line
'  hdr-ln  spec-item-hdr-line.  Non-Identent-Non-DD line
'Cml Catalog#Spec
'  Spec Specification
'  Tp Template
'
'Definition Spec
'  SpecTp
#End If
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbDtaLLn."
#If Doc Then
'Cml
' DDSRmk #Dash-Dash-Space-Rmk#
#End If
Type ILn: Ix As Integer: Ln As String: End Type ' Deriving(Ay Ctor)
Type Speci: Ix As Integer: Specit As String: Specin As String: Rst As String: ILny() As ILn: End Type 'Deriving(Ay Ctor)
Type Spec: IsLnMis As Boolean: IsSigMis As Boolean: Spect As String: Specn As String: IndSpec As String: Rmk() As String: Itms() As Speci: End Type 'Deriving(Ctor Opt)
Type SpecOpt: Som As Boolean: Spec As Spec: End Type

Function ILn(Ix, Ln) As ILn
With ILn
    .Ix = Ix
    .Ln = Ln
End With
End Function
Function AddILn(A As ILn, B As ILn) As ILn(): PushILn AddILn, A: PushILn AddILn, B: End Function
Sub PushILnAy(O() As ILn, A() As ILn): Dim J&: For J = 0 To ILnUB(A): PushILn O, A(J): Next: End Sub
Sub PushILn(O() As ILn, M As ILn): Dim N&: N = ILnSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function ILnSi&(A() As ILn): On Error Resume Next: ILnSi = UBound(A) + 1: End Function
Function ILnUB&(A() As ILn): ILnUB = ILnSi(A) - 1: End Function
Function Speci(Ix, Specit, Specin, Rst, ILny() As ILn) As Speci
With Speci
    .Ix = Ix
    .Specit = Specit
    .Specin = Specin
    .Rst = Rst
    .ILny = ILny
End With
End Function
Function AddSpeci(A As Speci, B As Speci) As Speci(): PushSpeci AddSpeci, A: PushSpeci AddSpeci, B: End Function
Sub PushSpeciAy(O() As Speci, A() As Speci): Dim J&: For J = 0 To SpeciUB(A): PushSpeci O, A(J): Next: End Sub
Sub PushSpeci(O() As Speci, M As Speci): Dim N&: N = SpeciSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SpeciSi&(A() As Speci): On Error Resume Next: SpeciSi = UBound(A) + 1: End Function
Function SpeciUB&(A() As Speci): SpeciUB = SpeciSi(A) - 1: End Function
Function Spec(Spect, Specn, IndSpec, Rmk$(), Itms() As Speci) As Spec
With Spec
    .Spect = Spect
    .Specn = Specn
    .IndSpec = IndSpec
    .Rmk = Rmk
    .Itms = Itms
End With
End Function
Function SpecOpt(Som, A As Spec) As SpecOpt: With SpecOpt: .Som = Som: .Spec = A: End With: End Function
Function SomSpec(A As Spec) As SpecOpt: SomSpec.Som = True: SomSpec.Spec = A: End Function
