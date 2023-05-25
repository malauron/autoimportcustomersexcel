
Imports MySql.Data.MySqlClient

Module modRoutines
    Public Function NetOpen(ByRef vData As DataTable, ByVal vSQL As String, Optional ByVal mConn As MySqlConnection = Nothing) As Boolean

        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter

        NetOpen = True

        vData = Nothing
        vData = New DataTable

        Try
            myCommand.Connection = mConn
            myCommand.CommandTimeout = 120
            myCommand.CommandText = vSQL
            myAdapter.SelectCommand = myCommand
            myAdapter.Fill(vData)
            myCommand.Dispose()
            myAdapter.Dispose()
        Catch myError As MySqlException
            Console.WriteLine("There was an error in executing your query: " & myError.Message)
            Return False
        End Try

    End Function
End Module
