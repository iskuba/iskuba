Imports System.Data.OleDb
Imports System.Data.SqlClient


Public Class Connection


    Public connectionString As String
    Public errorState As String
    Public errorMessage As String





    Function wykonajZapytanie(ByVal query As String)
        Try
        Dim da1 As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet

        da1 = New OleDbDataAdapter(query, connectionString)
        da1.Fill(ds, "query")
            Return ds.Tables("query")
            errorState = 0
        Catch ex As Exception
            Dim info As New info
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
                Dim info As New info
                info.RichTextBox1.Text = ex.Message
                info.ShowDialog()
                errorState = -1


            End If
            Return errorState
        End Try
    End Function




End Class
