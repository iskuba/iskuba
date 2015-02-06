Public Class menuERP

    Private login As login




    Public Sub setLogin(ByRef state As login)
        login = state
    End Sub

    Function returnLogin() As login
        Return login
    End Function



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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim forma As New ListaKontrahentow
        forma.setOknoMenu(Me)
        forma.ShowDialog()


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim forma As New listapracownikow
        forma.setOknoMenu(Me)
        forma.ShowDialog()
      
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim forma As New listatowarow
        forma.setOknoMenu(Me)
        forma.ShowDialog()

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim forma As New listaloginow
        forma.setOknoMenu(Me)
        forma.ShowDialog()
     

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim forma As New listadostaw
        forma.setOknoMenu(Me)
        forma.ShowDialog()

    End Sub
End Class