Public Class menuERP

    Private Sub menuERP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
End Class