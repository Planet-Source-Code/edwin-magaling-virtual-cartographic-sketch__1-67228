Attribute VB_Name = "modGlobal"
Public cartConn As ADODB.Connection
Public mainConn As ADODB.Connection
Public userConn As ADODB.Connection
Public mainRS As ADODB.Recordset
Public userRS As ADODB.Recordset
Public partsConn As ADODB.Connection
Public jawRS As ADODB.Recordset
Public hairRS As ADODB.Recordset
Public eyesRS As ADODB.Recordset
Public browRS As ADODB.Recordset
Public earsRS As ADODB.Recordset
Public noseRS As ADODB.Recordset
Public lipsRS As ADODB.Recordset
Public glassRS As ADODB.Recordset
Public beardRS As ADODB.Recordset
Public capRS As ADODB.Recordset
Public jawStr As String
Public hairStr As String
Public earsStr As String
Public browStr As String
Public eyesStr As String
Public noseStr As String
Public lipsStr As String
Public glassStr As String
Public beardStr As String
Public capStr As String
Public userCmd As ADODB.Command
Public mainCmd As ADODB.Command
Public userStr As String
Public mainStr As String

Sub main()
    frmLogin.Show
End Sub

Sub partsdbConnect()
    Set partsConn = New ADODB.Connection
    partsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\head_parts.mdb;Persist Security Info=False;Jet OLEDB:Database Password=winoe"
End Sub

Sub maindbConnect()
    Set mainConn = New ADODB.Connection
    mainConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\main.mdb;Persist Security Info=False;Jet OLEDB:Database Password=winoe"
End Sub

Sub userdbConnect()
    Set userConn = New ADODB.Connection
    userConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\users.mdb;Persist Security Info=False;Jet OLEDB:Database Password=winoe"
End Sub
