Public Class dodajedytujtowar


    Public id As Long


    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(TextBox1.Text).Length = 0 Then
            MsgBox("Podaj Kod  !")
            Exit Sub
        End If

        If Trim(TextBox2.Text).Length = 0 Then
            MsgBox("Podaj Nazwe!")
            Exit Sub
        End If



        If id = 0 Then '' jeżeli ID = 0 to oznacza ze tworzymy nowy rekord
            Form1.query.executeQuery("INSERT INTO twrkarty (twrkod,twrnazwa,jmz) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "')")
            MsgBox("Dodano Rekord !")
            listatowarow.ListaKontrahentow_Load(sender, e)
        Else
            Form1.query.executeQuery("UPDATE  twrKarty set twrkod= '" & TextBox1.Text & "',twrnazwa='" & TextBox2.Text & "',jmz='" & ComboBox1.Text & "' WHERE id=" & id & "")
            MsgBox("Rekord Został Zaktualizoany !")
            listatowarow.ListaKontrahentow_Load(sender, e)
        End If


    End Sub

   

    Private Sub dodajedytujkon_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        If id = 0 Then

            Me.Text = " Dodaj Towar"
        Else
            Me.Text = " Edytuj Towar"
        End If
    End Sub
End Class