Public Class dodajedytujdostawe


    Private oknoDostaw As listadostaw
    Private id As Long


    Public Sub setId(ByVal state As Long)
        id = state
    End Sub

    Public Sub setOknoTowar(ByRef state As listadostaw)
        oknoDostaw = state
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(TextBox1.Text).Length = 0 Then
            MsgBox("Podaj ilosc !")
            Exit Sub
        End If

      



        If id = 0 Then '' jeżeli ID = 0 to oznacza ze tworzymy nowy rekord

            oknoDostaw.returnOknoMenu.returnLogin.returnQuery.executeQuery("INSERT INTO traelem (idtwr,ilosc,iddost) VALUES (" & ComboBox1.SelectedValue & "," & TextBox1.Text & "," & ComboBox2.SelectedValue & ")")
            MsgBox("Dodano Rekord !")
            listadostaw.ListaKontrahentow_Load(sender, e)
        Else
            oknoDostaw.returnOknoMenu.returnLogin.returnQuery.executeQuery("UPDATE traelem set idtwr= " & ComboBox1.SelectedValue & ",ilosc=" & TextBox1.Text & ",iddost=" & ComboBox2.SelectedValue & " WHERE id=" & id & "")
            MsgBox("Rekord Został Zaktualizoany !")
            listadostaw.ListaKontrahentow_Load(sender, e)
        End If


    End Sub



    Private Sub dodajedytujkon_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim wynik As Object

        wynik = oknoDostaw.returnOknoMenu.returnLogin.returnQuery.wykonajZapytanie("SELECT id, twrkod  FROM twrkarty")

        If wynik.GetType.FullName = GetType(DataTable).FullName Then

            Dim tabela As DataTable

            tabela = wynik

            If tabela.Rows.Count > 0 Then
                ComboBox1.DataSource = tabela
                ComboBox1.DisplayMember = "twrkod"
                ComboBox1.ValueMember = "id"


            Else
                MsgBox("Brak zdefiniowanych towarów !")
                Exit Sub

            End If


        End If

        wynik = oknoDostaw.returnOknoMenu.returnLogin.returnQuery.wykonajZapytanie("SELECT id, nazwa  FROM kntkarty")

        If wynik.GetType.FullName = GetType(DataTable).FullName Then

            Dim tabela As DataTable

            tabela = wynik

            If tabela.Rows.Count > 0 Then
                ComboBox2.DataSource = tabela
                ComboBox2.DisplayMember = "nazwa"
                ComboBox2.ValueMember = "id"


            Else
                MsgBox("Brak zdefiniowanych dostawców !")
                Exit Sub

            End If


        End If



        ComboBox1.SelectedIndex = 0
        If id = 0 Then

            Me.Text = " Dodaj dostawce"
        Else
            Me.Text = " Edytuj dostawce"
        End If
    End Sub
End Class