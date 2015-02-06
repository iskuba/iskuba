Public Class dodajedytujtowar


    Private oknoTowar As listatowarow
    Private id As Long


    Public Sub setId(ByVal state As Long)
        id = state
    End Sub

    Public Sub setOknoTowar(ByRef state As listatowarow)
        oknoTowar = state
    End Sub




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

            oknoTowar.returnOknoMenu.returnLogin.returnQuery.executeQuery("INSERT INTO twrkarty (twrkod,twrnazwa,jmz) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "')")
            MsgBox("Dodano Rekord !")
            oknoTowar.ListaKontrahentow_Load(sender, e)
        Else
            oknoTowar.returnOknoMenu.returnLogin.returnQuery.executeQuery("UPDATE  twrKarty set twrkod= '" & TextBox1.Text & "',twrnazwa='" & TextBox2.Text & "',jmz='" & ComboBox1.Text & "' WHERE id=" & id & "")
            MsgBox("Rekord Został Zaktualizoany !")
            oknoTowar.ListaKontrahentow_Load(sender, e)
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