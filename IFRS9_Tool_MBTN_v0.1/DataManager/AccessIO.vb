Public Class AccessIO
    Private Connection As OleDb.OleDbConnection
    Private DataFile As String
    Private cmd As OleDb.OleDbCommand
    Private Adapter As OleDb.OleDbDataAdapter
    Public Reader As OleDb.OleDbDataReader

    Public Sub New()

    End Sub

    Public Sub New(DataFile As String)
        Me.DataFile = DataFile
    End Sub

    Public Sub OpenConnection()
        Connection = New OleDb.OleDbConnection
        Dim dbProvider As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="

        Dim connString As String = dbProvider & DataFile
        Connection.ConnectionString = connString
        Connection.Open()
    End Sub

    ' Create Command
    Public Sub CreateCommand(Query As String)
        cmd = New OleDb.OleDbCommand(Query, Connection)
    End Sub

    ' Execute NonQuery
    Public Sub ExecuteNonQuery()
        cmd.ExecuteNonQuery()
    End Sub

    ' Execute Reader
    Public Sub ExecuteReader()
        Reader = cmd.ExecuteReader()
    End Sub

    ' Close Reader
    Public Sub CloseReader()
        Reader.Close()
    End Sub

    ' Close Connection
    Public Sub CloseConnection()
        Connection.Close()
    End Sub

    Public Sub CheckConnectionStatus()
        CheckConnectionStatus()
    End Sub

End Class
