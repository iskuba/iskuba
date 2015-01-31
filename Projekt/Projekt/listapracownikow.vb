Public Class listapracownikow

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
            wynik = login.query.wykonajZapytanie("SELECT * FROM prckarty")
            If wynik.GetType.FullName = GetType(DataTable).FullName Then
                DataGridView1.DataSource = wynik
                DataGridView1.Columns("idprc").Visible = False
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

        table2.DefaultView.RowFilter = "imie like '" & TextBox1.Text & "%' and nazwisko like '" & TextBox2.Text & "%' and miejscowosc like '" & TextBox3.Text & "%' "










    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        sortuj()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        sortuj()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dodajedytujprc.Close()
        dodajedytujprc.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If IsNothing(DataGridView1.CurrentRow) Then

        Else
            On Error Resume Next
            dodajedytujprc.Close()
            dodajedytujprc.id = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            dodajedytujprc.TextBox1.Text = DataGridView1.Item("imie", DataGridView1.CurrentRow.Index).Value
            dodajedytujprc.TextBox2.Text = DataGridView1.Item("miejscowosc", DataGridView1.CurrentRow.Index).Value
            Dim kod As String

            kod = DataGridView1.Item("kodpocztowy", DataGridView1.CurrentRow.Index).Value
            dodajedytujprc.TextBox3.Text = kod(0)
            dodajedytujprc.TextBox4.Text = kod(1)
            dodajedytujprc.TextBox5.Text = DataGridView1.Item("telefon", DataGridView1.CurrentRow.Index).Value
            dodajedytujprc.TextBox6.Text = DataGridView1.Item("nrlokalu", DataGridView1.CurrentRow.Index).Value
            dodajedytujprc.TextBox7.Text = DataGridView1.Item("ulica", DataGridView1.CurrentRow.Index).Value
            dodajedytujprc.TextBox8.Text = DataGridView1.Item("nazwisko", DataGridView1.CurrentRow.Index).Value
            dodajedytujprc.Show()

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try

            If IsNothing(DataGridView1.CurrentRow) Then

            Else

                Dim result As Integer

                result = MsgBox("Czy na pewno chcesz usunąć pracownika ?", vbYesNo)

                If result = 6 Then


                Else
                    Exit Sub
                End If


                Dim wynik As Object
                wynik = login.query.wykonajZapytanie("SELECT idprc FROM  usr where idprc =" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & " UNION ALL SELECT idpracownika FROM TraNag Where idpracownika =" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")

                If wynik.GetType.FullName = GetType(DataTable).FullName Then
                    Dim tabela As DataTable

                    tabela = wynik

                    If tabela.Rows.Count > 0 Then

                        MsgBox("Istnieją wiązania z wybranym pracownikiem. Usuwanie rekordu zostanie anulowane !")
                    Else

                        login.query.executeQuery("DELETE FROM prckarty where idprc=" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")
                        Me.ListaKontrahentow_Load(sender, e)
                    End If


                End If

            End If
        Catch ex As Exception
            dodajedytujkon.Show()
        End Try
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        sortuj()

    End Sub
End Class