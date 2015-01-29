Public Class dodajedytujkon

    Public id As Long


    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(TextBox1.Text).Length = 0 Then
            MsgBox("Podaj Nazwę !")
            Exit Sub
        End If

        If Trim(TextBox2.Text).Length = 0 Or Trim(TextBox3.Text).Length = 0 Or Trim(TextBox4.Text).Length = 0 Or Trim(TextBox6.Text).Length = 0 Or Trim(TextBox7.Text).Length = 0 Then
            MsgBox("Podaj Dane Adresowe (Miejscowość, Kod Pocztowy , Numer Lokalu) !")
            Exit Sub
        End If



        If id = 0 Then '' jeżeli ID = 0 to oznacza ze tworzymy nowy rekord
            Form1.query.executeQuery("INSERT INTO kntKarty (nazwa,miejscowosc,kodpocztowy,telefon,nrlokalu,ulica) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "-" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "')")
            MsgBox("Dodano Rekord !")
            ListaKontrahentow.ListaKontrahentow_Load(sender, e)
        Else
            Form1.query.executeQuery("UPDATE  kntKarty set nazwa= '" & TextBox1.Text & "',miejscowosc='" & TextBox2.Text & "',kodpocztowy='" & TextBox3.Text & "-" & TextBox4.Text & "',telefon='" & TextBox5.Text & "',nrlokalu='" & TextBox6.Text & "',ulica='" & TextBox7.Text & "' WHERE id=" & id & "")
            MsgBox("Rekord Został Zaktualizoany !")
            ListaKontrahentow.ListaKontrahentow_Load(sender, e)
        End If


    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text.Length > 0 Then
            If IsNumeric(TextBox3.Text) Then
            Else
                TextBox3.Text = ""
            End If
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text.Length > 0 Then
            If IsNumeric(TextBox4.Text) Then
            Else
                TextBox4.Text = ""
            End If
        End If
    End Sub

    Private Sub dodajedytujkon_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If id = 0 Then

            Me.Text = " Dodaj Kontrahenta "
        Else
            Me.Text = " Edytuj Kontrahenta "
        End If
    End Sub
End Class