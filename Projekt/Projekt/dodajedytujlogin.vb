Public Class dodajedytujlogin

    Private oknoLog As listaloginow
    Private id As Long


    Public Sub setId(ByVal state As Long)
        id = state
    End Sub

    Public Sub setOknoTowar(ByRef state As listaloginow)
        oknoLog = state
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(TextBox1.Text).Length = 0 Then
            MsgBox("Akronim  !")
            Exit Sub
        End If

        If Trim(MaskedTextBox1.Text).Length = 0 Then
            MsgBox("Podaj hasło!")
            Exit Sub
        End If


        Dim admin As Integer

        If CheckBox1.Checked Then

            admin = 1
        Else
            admin = 0

        End If


        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(MaskedTextBox1.Text)
        Dim haslo As String


        haslo = System.Convert.ToBase64String(byt)



        If id = 0 Then '' jeżeli ID = 0 to oznacza ze tworzymy nowy rekord
            oknoLog.returnOknoMenu.returnLogin.returnQuery.executeQuery("INSERT INTO usr (akronim,haslo,admin,idprc) VALUES ('" & TextBox1.Text & "','" & haslo & "'," & admin & "," & ComboBox1.SelectedValue & ")")
            MsgBox("Dodano Rekord !")
            oknoLog.ListaKontrahentow_Load(sender, e)
        Else
            oknoLog.returnOknoMenu.returnLogin.returnQuery.executeQuery("UPDATE  usr set akronim= '" & TextBox1.Text & "',haslo='" & haslo & "',admin=" & admin & ", idprc=" & ComboBox1.SelectedValue & " WHERE id_usr=" & id & "")
            MsgBox("Rekord Został Zaktualizoany !")
            oknoLog.ListaKontrahentow_Load(sender, e)
        End If


    End Sub



    Private Sub dodajedytujkon_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim wynik As Object

        wynik = oknoLog.returnOknoMenu.returnLogin.returnQuery.wykonajZapytanie("SELECT idprc, imie + ' ' + nazwisko as nazwa FROM prckarty")

        If wynik.GetType.FullName = GetType(DataTable).FullName Then

            Dim tabela As DataTable

            tabela = wynik

            If tabela.Rows.Count > 0 Then
                ComboBox1.DataSource = tabela
                ComboBox1.DisplayMember = "nazwa"
                ComboBox1.ValueMember = "idprc"


            Else
                MsgBox("Brak zdefiniowanych pracowników !")
                Exit Sub

            End If





        End If

        ComboBox1.SelectedIndex = 0
        If id = 0 Then

            Me.Text = " Dodaj Login"
        Else
            Me.Text = " Edytuj Login"
        End If
    End Sub
End Class