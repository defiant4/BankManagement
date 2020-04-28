Attribute VB_Name = "Module1"

Option Explicit
Public UserType As String
'Public userdelete As Boolean, edituser As Boolean
'Public recCustomer As New ADODB.Recordset
Public Cn As ADODB.Connection

Public Sub opencon()
Set Cn = New ADODB.Connection
Cn.CursorLocation = adUseClient
Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database1.mdb" & "; Persist Security Info=False;"
Cn.Open
End Sub

'CHECKING Username and password
Public Function validuser(Username As String, password As String, status As String)

Dim db As String
Dim Cmd As String
Dim sql As String
Dim Cn As ADODB.Connection
Dim rs As ADODB.Recordset

db = App.Path & "\database1.MDB"
Cmd = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & db & ""

Set Cn = New ADODB.Connection

With Cn
    .ConnectionString = Cmd
    .Open
End With

Set rs = New ADODB.Recordset


sql = "Select * From [login] where Username LIKE '" & Username & "' and Password LIKE '" & password & "' and Status LIKE '" & status & "'"
rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

 
 If Not rs.EOF Then
 validuser = True
 Else
 validuser = False
 End If
 
 rs.Close
 Set rs = Nothing

 Cn.Close
 Set Cn = Nothing
End Function
'AddUser
Public Function AddUser(Username As String, password As String, status As String)

Dim db As String
Dim Cmd As String
Dim sql As String
Dim Cn As ADODB.Connection
Dim rs As ADODB.Recordset

db = App.Path & "\database1.MDB"
Cmd = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & db & ""

Set Cn = New ADODB.Connection

With Cn
    .ConnectionString = Cmd
    .Open
End With

Set rs = New ADODB.Recordset


sql = "Select * From [login]"
rs.Open sql, Cn, adOpenForwardOnly, adLockOptimistic

 
 If Not rs.EOF Then
 With rs
 .AddNew
 rs("Username") = Username
 rs("Password") = password
 rs("Status") = status
 .update
End With
 AddUser = True
 Else
 AddUser = False
 End If
 
 rs.Close
 Set rs = Nothing

 Cn.Close
 Set Cn = Nothing
End Function

'CHECKING EXISTING User
Public Function CHECKUSER(USER As String)
Dim db As String
Dim Cmd As String
Dim sql As String
Dim Cn As ADODB.Connection
Dim rs As ADODB.Recordset

db = App.Path & "\database1.MDB"
Cmd = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & db & ""

Set Cn = New ADODB.Connection

With Cn
    .ConnectionString = Cmd
    .Open
End With

Set rs = New ADODB.Recordset


sql = "Select * From [login] where Username LIKE '" & USER & "'"
rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

 
 If Not rs.EOF Then
 CHECKUSER = True
 Else
 CHECKUSER = False
 End If
 
 rs.Close
 Set rs = Nothing

 Cn.Close
 Set Cn = Nothing
End Function

