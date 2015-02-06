''' <summary>
''' Klasa tworząca menu naszego systemu
''' </summary>



Public Class menuERP


    ''' <summary>
    ''' Zmienna przecowująca referencję do głównej klasy programu.
    ''' </summary>
    Private login As login



    ''' <summary>
    ''' Metoda ustawiająca zmieną login
    ''' </summary>
    Public Sub setLogin(ByRef state As login)
        login = state
    End Sub
    ''' <summary>
    ''' Metoda zwracająca zmieną login
    ''' </summary>
    Function returnLogin() As login
        Return login
    End Function


    ''' <summary>
    ''' Metoda ładująca okno formy
    ''' </summary>
    Private Sub menuERP_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        If (login.CurrentUser.returnIsAdmin() = 1) Then
            ToolStripButton1.Visible = True
        Else
            ToolStripButton1.Visible = False
        End If

        Dim wynik As Object



        wynik = login.returnQuery().wykonajZapytanie("SELECT imie + ' ' + Nazwisko FROM prckarty where idprc=" & login.CurrentUser().returnIdUser & "")

        If wynik.GetType.FullName = GetType(DataTable).FullName Then
            Dim tabela As DataTable
            tabela = wynik

            If tabela.Rows.Count > 0 Then
                ToolStripLabel1.Text = "Zalogowano przez : " & tabela.Rows(0).Item(0)
            Else
                MsgBox("Brak takiego użytkownika w bazie !")
                Me.Close()
            End If


        Else
            Me.Close()
        End If

    End Sub


    ''' <summary>
    ''' Metoda wywołująca button2Click, tworzny obiekt typu ListaKontrahentow oraz wywołuje metodę ładującą forme klasy
    ''' </summary>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim forma As New ListaKontrahentow
        forma.setOknoMenu(Me)
        forma.ShowDialog()


    End Sub

    ''' <summary>
    ''' Metoda wywołująca button4Click, tworzny obiekt typu ListaPracowników oraz wywołuje metodę ładującą forme klasy
    ''' </summary>
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim forma As New listapracownikow
        forma.setOknoMenu(Me)
        forma.ShowDialog()

    End Sub

    ''' <summary>
    ''' Metoda wywołująca button1Click, tworzny obiekt typu ListaTowarow oraz wywołuje metodę ładującą forme klasy
    ''' </summary>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim forma As New listatowarow
        forma.setOknoMenu(Me)
        forma.ShowDialog()

    End Sub

    ''' <summary>
    ''' Metoda wywołująca ToolStripButton1Click, tworzny obiekt typu ListaLoginów oraz wywołuje metodę ładującą forme klasy
    ''' </summary>
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim forma As New listaloginow
        forma.setOknoMenu(Me)
        forma.ShowDialog()


    End Sub
    ''' <summary>
    ''' Metoda wywołująca button3Click, tworzny obiekt typu ListaDostaw oraz wywołuje metodę ładującą forme klasy
    ''' </summary>
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim forma As New listadostaw
        forma.setOknoMenu(Me)
        forma.ShowDialog()

    End Sub
End Class