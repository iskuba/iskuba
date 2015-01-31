Public Class ListaKontrahentow

    Public Sub ListaKontrahentow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Me.SetStyle(
ControlStyles.AllPaintingInWmPaint Or _
ControlStyles.UserPaint Or _
ControlStyles.DoubleBuffer, True)


            Dim p = GetType(DataGridView).GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            p.SetValue(Me.DataGridView1, True, Nothing)



            DataGridView1.AllowUserToAddRows = False
            DataGridView1.AllowUserToDeleteRows = False
            DataGridView1.AllowUserToResizeRows = False
            DataGridView1.EnableHeadersVisualStyles = False
            DataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect
            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke
            DataGridView1.ShowEditingIcon = False

            DataGridView1.TabIndex = 0
            DataGridView1.RowHeadersWidth = 55
            DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            DataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            DataGridView1.RowHeadersVisible = False



            Dim wynik As Object
            wynik = login.query.wykonajZapytanie("SELECT * FROM kntkarty")
            If wynik.GetType.FullName = GetType(DataTable).FullName Then
                DataGridView1.DataSource = wynik
                DataGridView1.Columns("id").Visible = False
            End If

            sortuj()

        Catch ex As Exception
            Dim info As New info
            info.RichTextBox1.Text = ex.Message
            info.ShowDialog()
            Me.Close()
        End Try
    End Sub


    Sub sortuj()
        Dim table2 As DataTable
        table2 = DataGridView1.DataSource




        '' MsgBox("Kod_Towaru " & kod & " and Numer_Zlecenia " & zw & "")
        ''    MsgBox("Twr_Kod like '" & TextBox1.Text & "%' and NumerZlecenia like '" & TextBox2.Text & "%'and Dostawa like '" & TextBox3.Text & "%' and NumerPrzewodnika like '" & TextBox4.Text & "%' and  " & rw & "")

        table2.DefaultView.RowFilter = "nazwa like '" & TextBox1.Text & "%' and miejscowosc like '" & TextBox2.Text & "%'"










    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        sortuj()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        sortuj()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dodajedytujkon.Close()
        dodajedytujkon.Show()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If IsNothing(DataGridView1.CurrentRow) Then

        Else
            On Error Resume Next
            dodajedytujkon.Close()
            dodajedytujkon.id = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            dodajedytujkon.TextBox1.Text = DataGridView1.Item("nazwa", DataGridView1.CurrentRow.Index).Value
            dodajedytujkon.TextBox2.Text = DataGridView1.Item("miejscowosc", DataGridView1.CurrentRow.Index).Value
            Dim kod As String

            kod = DataGridView1.Item("kodpocztowy", DataGridView1.CurrentRow.Index).Value
            dodajedytujkon.TextBox3.Text = kod(0)
            dodajedytujkon.TextBox4.Text = kod(1)
            dodajedytujkon.TextBox5.Text = DataGridView1.Item("telefon", DataGridView1.CurrentRow.Index).Value
            dodajedytujkon.TextBox6.Text = DataGridView1.Item("nrlokalu", DataGridView1.CurrentRow.Index).Value
            dodajedytujkon.TextBox7.Text = DataGridView1.Item("ulica", DataGridView1.CurrentRow.Index).Value
            dodajedytujkon.Show()

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try

            If IsNothing(DataGridView1.CurrentRow) Then

            Else

                Dim result As Integer

                result = MsgBox("Czy na pewno chcesz usunąć kontrahenta ?", vbYesNo)

                If result = 6 Then


                Else
                    Exit Sub
                End If


                Dim wynik As Object
                wynik = login.query.wykonajZapytanie("SELECT * FROM  tranag where idkontrahenta=" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")

                If wynik.GetType.FullName = GetType(DataTable).FullName Then
                    Dim tabela As DataTable

                    tabela = wynik

                    If tabela.Rows.Count > 0 Then

                        MsgBox("Istnieją tranzakcje na tego kontrahenta. Usuwanie rekordu zostanie anulowane !")
                    Else

                        login.query.executeQuery("DELETE FROM KNTKARTY where id=" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")
                        Me.ListaKontrahentow_Load(sender, e)
                    End If


                End If

            End If
        Catch ex As Exception
            dodajedytujkon.Show()
        End Try
    End Sub
End Class