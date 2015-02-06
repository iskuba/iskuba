
''' <summary>
''' Klasa ktora odtwarzana jest na poczatku aplikacji.  Dzieki tej klasie logujemy
''' sie do systemu ERP
''' </summary>
Public Class login

    ''' <summary>
    ''' Zmienna typu Connection.
    ''' </summary>
    Private query As New Connection
    ''' <summary>
    ''' Zmienna typu currentUser.
    ''' </summary>
    Private user As New currentUser

    ''' <summary>
    ''' Metoda ładująca okno klasy
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        query.SetConnectionString( _
    "Provider=Microsoft.Jet.OLEDB.4.0;User Id= zbud; Password=z; Data source=" & _
    "" & Application.StartupPath & "\baza\bazaERP.mdb; Jet OLEDB:System Database=" & Application.StartupPath & "\baza\Zabezpieczenia4.mdw;User ID=Paweł;Password=;")

    End Sub


    ''' <summary>
    ''' Metoda zwracająca przez referencję zmieną query
    ''' </summary>
    Function returnQuery() As Connection
        Return query
    End Function
    ''' <summary>
    ''' 
    ''' Metoda zwracająca przez referencję zmieną user
    ''' </summary>
    Function CurrentUser() As currentUser
        Return user
    End Function


    ''' <summary>
    ''' Metoda wywołująca reakcję na buttonie, wywołuje logowanie do systemu ERP
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Trim(TextBox1.Text).Length = 0 Then
            MsgBox("Nie Podano Loginu!")
            Exit Sub
        End If

        If Trim(MaskedTextBox1.Text).Length = 0 Then
            MsgBox("Nie Podano Hasła!")
            Exit Sub
        End If


        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(MaskedTextBox1.Text)
        Dim haslo As String


        haslo = System.Convert.ToBase64String(byt)

        Dim wynik As Object

        wynik = query.wykonajZapytanie("SELECT idprc,admin from usr WHERE haslo ='" & haslo & "' and akronim='" & TextBox1.Text & "'")

        If wynik.GetType.FullName = GetType(DataTable).FullName Then
            Dim tabela As DataTable
            tabela = wynik

            If tabela.Rows.Count > 0 Then
                user.setIDusr(tabela.Rows(0).Item("idprc"))
                user.setIsAdmin(tabela.Rows(0).Item("admin"))
                Dim forma As New menuERP
                forma.setLogin(Me)
                forma.ShowDialog()
            Else
                user = Nothing
                MsgBox("Podane dane są błędne!")
                Exit Sub
            End If
        End If

    End Sub
End Class
