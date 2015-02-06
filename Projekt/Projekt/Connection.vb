Imports System.Data.OleDb
Imports System.Data.SqlClient


Public Class Connection


    Private connectionString As String
    Private errorState As String
    Private errorMessage As String
    Private info As New info







    Public Sub SetConnectionString(ByVal connection As String)
        connectionString = connection
    End Sub

    Function wykonajZapytanie(ByVal query As String)
        Try
            Dim da1 As New System.Data.OleDb.OleDbDataAdapter
            Dim ds As New DataSet

            da1 = New OleDbDataAdapter(query, connectionString)
            da1.Fill(ds, "query")
            Return ds.Tables("query")
            errorState = 0
        Catch ex As Exception
            info.RichTextBox1.Text = ex.Message
            info.ShowDialog()
            errorState = -1
            Return errorState
        End Try
    End Function


    Function executeQuery(ByVal query As String)
        Try
            Dim MyConnection As New OleDbConnection(connectionString)
            Dim command1 As New OleDbCommand(query, MyConnection)
            command1.Connection.Open()




            errorState = command1.ExecuteScalar()


            MyConnection.Close()

            Return errorState
        Catch ex As Exception


            If ex.Message Like "Conversion from type 'DBNull' to type 'String' is not valid." Then


                errorState = 0
            Else
                info.RichTextBox1.Text = ex.Message
                info.ShowDialog()
                errorState = -1
            End If
            Return errorState
        End Try
    End Function

End Class
