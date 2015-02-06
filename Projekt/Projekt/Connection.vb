Imports System.Data.OleDb
Imports System.Data.SqlClient

''' <summary>
''' Klasa dzięku której łączymy się z bazą sql
''' </summary>
Public Class Connection

    ''' <summary>
    ''' Zmiena przechowująca adres do połączenia z bazą
    ''' </summary>
    Private connectionString As String
    ''' <summary>
    ''' Zmienna przechowująca błąd w postaci numeru błędu który powstał podczas połączenia,edycji w bazie
    ''' </summary>
    Private errorState As String
    ''' <summary>
    ''' Zmienna przechowująca błąd w postaci tekstu błędu który powstał podczas połączenia,edycji w bazie
    ''' </summary>
    Private errorMessage As String
    ''' <summary>
    ''' Zmienna przechowująca referencję do klasy info, która wyswietla błędy
    ''' </summary>
    Private info As New info






    ''' <summary>
    ''' Metoda ładująca wartosc do connectionString
    ''' </summary>
    Public Sub SetConnectionString(ByVal connection As String)
        connectionString = connection
    End Sub
    ''' <summary>
    ''' Metoda wykonująca zapytania typu SELECT
    ''' </summary>
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

    ''' <summary>
    ''' Metoda wykonująca zapytania typu DELETE,UPDATE,INSERT 
    ''' </summary>
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
