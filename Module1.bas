Attribute VB_Name = "Module1"
Public ConN As New ADODB.Connection
Public SdR As New ADODB.Recordset
Public MySql As String
Public SrM As ADODB.Stream
Public CmD As ADODB.Command

Public Sub cNDb()
Set ConN = New ADODB.Connection
Set SdR = New ADODB.Recordset

ConN.Open "Driver={Mysql ODBC 5.2 UNICODE Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=kenaikan_jabatan;"
ConN.CursorLocation = adUseClient
End Sub
