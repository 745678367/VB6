Attribute VB_Name = "ConnectAccess2"
Private JRO As New JetEngine

'连接数据库
Public Sub ConnectAccess(DBcnn As ADODB.Connection, DBset As ADODB.Recordset)
    Set DBcnn = New ADODB.Connection
    DBcnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB.mdb" & ";Persist Security Info=False"
    DBcnn.Open
    
    Set DBset = New ADODB.Recordset
    With DBset
        .CursorType = adOpenDynamic
        .LockType = adLockPessimistic
        .CursorLocation = adUseClient
    End With
End Sub

'刷新数据库
Public Sub RefreshDB(DBcnn As ADODB.Connection)
    JRO.RefreshCache DBcnn
End Sub
