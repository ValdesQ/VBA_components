VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabaseConnention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private oConnectionDictionayTypes As Dictionary
Private oConnectionDictionayParameters As Dictionary
Private oConnection As Object

Enum DatabaseTypeConnection
    OLEDB_SqlServer = 0
    OLEDB_Access = 1
    ODBC_Text_standart = 2
    ODBC_Text_64 = 3
    ODBC_Excel_old = 4
    ODBC_Excel_new = 5
    ODBC_Excel_macro = 6
    Jet_Excel_old = 7
    Jet_Text = 8
    
End Enum

Private Function fncCreateConnectionString(ByVal sConnectionString As String) As String
    Dim oKey As Variant
    
    
    For Each oKey In oConnectionDictionayParameters.Keys
        sConnectionString = Replace(sConnectionString, oKey, oConnectionDictionayParameters(oKey))
    Next oKey
    
    fncCreateConnectionString = sConnectionString
    
End Function

Private Sub prcConnectionDictionayTypes()
    Set oConnectionDictionayTypes = New Scripting.Dictionary
    
    oConnectionDictionayTypes(0) = Array("OLEDB_SqlServer", "Provider=sqloledb;Data Source={Source} ;Initial Catalog={Catalog};User Id={User};Password={Pass};")
    oConnectionDictionayTypes(1) = Array("OLEDB_Access", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Source};Persist Security Info=False;")
    oConnectionDictionayTypes(2) = Array("ODBC_Text_standart", "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq={Source};Extensions=asc,csv,tab,txt;")
    oConnectionDictionayTypes(3) = Array("ODBC_Text_64", "Driver=Microsoft Access Text Driver (*.txt, *.csv);Dbq={Source};Extensions=asc,csv,tab,txt;")
    oConnectionDictionayTypes(4) = Array("ODBC_Excel_old", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Source};Extended Properties=""Excel 8.0;HDR=YES"";")
    oConnectionDictionayTypes(5) = Array("ODBC_Excel_new", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Source};Extended Properties=""Excel 12.0 Xml;HDR=YES"";")
    oConnectionDictionayTypes(6) = Array("ODBC_Excel_macro", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Source};Extended Properties=""Excel 12.0 Macro;HDR=YES"";")
    oConnectionDictionayTypes(7) = Array("Jet_Excel_old", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={Source};Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";")
    oConnectionDictionayTypes(8) = Array("Jet_Text", "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq={Source};Extensions=asc,csv,tab,txt;")

End Sub


Public Sub OpenConnection( _
    ByVal sType As DatabaseTypeConnection, _
    Optional ByVal sServer As String = vbNullString, _
    Optional ByVal sDatabase As String = vbNullString, _
    Optional ByVal sUser As String = vbNullString, _
    Optional ByVal sPass As String = vbNullString _
)
    Call prcConnectionDictionayTypes
    
    Set oConnectionDictionayParameters = New Scripting.Dictionary
    oConnectionDictionayParameters("{Source}") = sServer
    oConnectionDictionayParameters("{Catalog}") = sDatabase
    oConnectionDictionayParameters("{User}") = sUser
    oConnectionDictionayParameters("{Pass}") = sPass
    
    Set oConnection = CreateObject("ADODB.connection")
    oConnection.ConnectionTimeout = 120
    
    oConnection.connectionString = fncCreateConnectionString(oConnectionDictionayTypes(sType)(1))
    oConnection.Open
    
End Sub


Public Sub ExecuteSQL(ByVal sqlString As String)
    oConnection.Execute (sqlString)

End Sub
Public Function SqlResult(ByVal sqlString As String) As Object
    Set SqlResult = oConnection.Execute(sqlString)

End Function


Private Sub Class_Terminate()
    oConnection.Close
    Set oConnection = Nothing
    
End Sub
