Attribute VB_Name = "MxDtaDaSamp"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaSamp."
'From:
'https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sql/sql-server-express-user-instances
Public Const SampCnStr_SQLEXPR$ = "Data Source=.\\SQLExpress;Integrated Security=true;" & _
"User Instance=true;AttachDBFilename=|DataDirectory|\InstanceDB.mdf;" & _
"Initial Catalog=InstanceDB;"
'------------------------------------------
'From:
'https://social.msdn.microsoft.com/Forums/vstudio/en-US/61d45bef-eea7-4366-a8ad-e15a1fa3d544/vb6-to-connect-with-sqlexpress?forum=vbgeneral
Public Const SampCnStr_SQLEXPR_NotWrk3$ = _
"Provider=SQLNCLI.1;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
Public Const GetCnStr_ADO_SampSQL_EXPR_NOT_WRK$ = _
"Provider=LoSqleDb;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
'--------------------------------
'From https://social.msdn.microsoft.com/Forums/en-US/a73a838b-ec3f-419b-be65-8b1732fbf4d0/connect-to-a-remote-sql-server-db?forum=isvvba
Public Const SampCnStr_SQLEXPR_NotWrk1$ = "driver={SQL Server};" & _
      "server=LAPTOP-SH6AEQSO;uid=MyUserName;pwd=;database=pubs"
   
Public Const SampCnStr_SQLEXPR_NotWrk2$ = "driver={SQL Server};" & _
      "server=127.0.0.1;uid=MyUserName;pwd=;database=pubs"
   
Public Const SampCnStr_SQLEXPR_NotWrk$ = ".\SQLExpress;AttachDbFilename=c:\mydbfile.mdf;Database=dbname;" & _
"Trusted_Connection=Yes;"
'"Typical normal SQL Server connection string: Data Source=myServerAddress;
'"Initial Catalog=myDataBase;Integrated Security=SSPI;"

'From VisualStudio
Public Const SampSqlCnStr_NotWrk$ = _
    "Data Source=LAPTOP-SH6AEQSO\ProjectsV13;Initial Catalog=master;Integrated Security=True;Connect Timeout=30;" & _
    "Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"
Property Get SampDr1() As Variant()
SampDr1 = Array(1, 2, 3)
End Property

Property Get SampDr2() As Variant()
SampDr2 = Array(2, 3, 4)
End Property

Property Get SampDr3() As Variant()
SampDr3 = Array(3, 4, 5)
End Property

Property Get SampDr4() As Variant()
SampDr4 = Array(43, 44, 45)
End Property

Property Get SampDr5() As Variant()
SampDr5 = Array(53, 54, 55)
End Property

Property Get SampDr6() As Variant()
SampDr6 = Array(63, 64, 65)
End Property

Property Get SampDrs2() As Drs
SampDrs2 = DrszFF("A B C", SampDy2)
End Property

Property Get SampDrs1() As Drs
SampDrs1 = DrszFF("A B C", SampDy1)
End Property

Property Get SampDrs() As Drs
SampDrs = DrszFF("A B C D E G H I J K", SampDy)
End Property

Property Get SampDy1() As Variant()
SampDy1 = Array(SampDr1, SampDr2, SampDr3)
End Property

Property Get SampDy2() As Variant()
SampDy2 = Array(SampDr3, SampDr4, SampDr5)
End Property

Property Get SampDy3() As Variant()
PushI SampDy3, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(100, "A") & vbCrLf & String(100, "X"))
PushI SampDy3, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(100, "A") & vbCrLf & String(100, "X"))
End Property

Property Get SampDy() As Variant()
PushI SampDy, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "A"))
PushI SampDy, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "B"))
PushI SampDy, Array("C", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "C"))
PushI SampDy, Array("D", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "D"))
PushI SampDy, Array("E", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "E"))
PushI SampDy, Array("F", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "F"))
PushI SampDy, Array("G", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "G"))
End Property

Function SampDs() As Ds
PushDt SampDs.DtAy, SampDt1
PushDt SampDs.DtAy, SampDt2
SampDs.DsNm = "Ds"
End Function

Property Get SampDt1() As Dt
SampDt1 = DtByFF("SampDt1", "A B C", SampDy1)
End Property

Property Get SampDt2() As Dt
SampDt2 = DtByFF("SampDt2", "A B C", SampDy2)
End Property

Property Get SampDr_AToJ() As Variant()
Const NC% = 10
Dim J%
For J = 0 To NC - 1
    PushI SampDr_AToJ, Chr(Asc("A") + J)
Next
End Property
