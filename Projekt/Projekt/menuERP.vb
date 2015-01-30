Public Class menuERP

    Private Sub menuERP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If (Form1.user.isadmin = 1) Then
            ToolStripButton1.Visible = True
        Else
            ToolStripButton1.Visible = False
        End If

        Dim wynik As Object

        wynik = Form1.query.wykonajZapytanie("SELECT imie + ' ' + Nazwisko FROM prckarty where idprc=" & Form1.user.idusr & "")

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
        ListaKontrahentow.ShowDialog()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        listapracownikow.ShowDialog()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        listatowarow.ShowDialog()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        listaloginow.Close()
        listaloginow.Show()

    End Sub
End Class