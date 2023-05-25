Imports MySql.Data.MySqlClient

Public NotInheritable Class clsConnectionDetails

    Private vHost As String
    Private vUsername As String
    Private vPassword As String
    Private vPort As String
    Private vDatabase As String
    Private vException As String

    Private vIsConnected As Boolean

    Private vConnection As MySqlConnection

    Public Function Host() As String
        Return vHost
    End Function

    Public Function Username() As String
        Return vUsername
    End Function

    Public Function Password() As String
        Return vPassword
    End Function

    Public Function Port() As String
        Return vPort
    End Function

    Public Function Database() As String
        Return vDatabase
    End Function

    Public Function IsConnected() As Boolean
        Return vIsConnected
    End Function

    Public Function Exception() As String
        Return vException
    End Function

    Public Function Connection() As MySqlConnection
        Return vConnection
    End Function

    Public Sub ConnectToServer(ByVal mHost As String, ByVal mUsername As String, ByVal mPassword As String, ByVal mPort As String, ByVal mDatabase As String)

        vConnection = Nothing
        vConnection = New MySqlConnection

        vHost = mHost
        vUsername = mUsername
        vPassword = mPassword
        vPort = mPort
        vDatabase = mDatabase

        vIsConnected = False

        vConnection.ConnectionString = "Server=" & vHost & ";" _
                & "Uid=" & vUsername & ";" _
                & "Pwd=" & vPassword & ";" _
                & "Port=" & vPort & ";" _
                & "database=" & vDatabase & ";" _
                & "Allow User Variables=True"

        Try
            vConnection.Open()
            vIsConnected = True
        Catch myerror As MySqlException
            vException = myerror.Message
        Catch myerror As Exception
            vException = myerror.Message
        End Try

    End Sub

End Class
