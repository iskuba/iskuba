Public Class listaloginow

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
            wynik = Form1.query.wykonajZapytanie("SELECT * FROM usr")
            If wynik.GetType.FullName = GetType(DataTable).FullName Then
                DataGridView1.DataSource = wynik
                DataGridView1.Columns("id_usr").Visible = False
                DataGridView1.Columns("haslo").Visible = False
                DataGridView1.Columns("idprc").Visible = False
                DataGridView1.Columns("admin").Visible = False
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

        table2.DefaultView.RowFilter = "akronim like '" & TextBox1.Text & "%' "










    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        sortuj()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs)
        sortuj()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dodajedytujlogin.Close()
        dodajedytujlogin.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If IsNothing(DataGridView1.CurrentRow) Then

        Else
            On Error Resume Next
            dodajedytujprc.Close()
            dodajedytujlogin.id = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            dodajedytujlogin.TextBox1.Text = DataGridView1.Item("akronim", DataGridView1.CurrentRow.Index).Value

            Dim haslo As Byte()
            haslo = System.Convert.FromBase64String(DataGridView1.Item("haslo", DataGridView1.CurrentRow.Index).Value)
            dodajedytujlogin.MaskedTextBox1.Text = System.Text.Encoding.UTF8.GetString(haslo)

            If DataGridView1.Item("admin", DataGridView1.CurrentRow.Index).Value = 1 Then
                dodajedytujlogin.CheckBox1.Checked = True
            Else
                dodajedytujlogin.CheckBox1.Checked = False
            End If
            dodajedytujlogin.Show()

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try

            If IsNothing(DataGridView1.CurrentRow) Then

            Else

                Dim result As Integer

                result = MsgBox("Czy na pewno chcesz usunąć login ?", vbYesNo)

                If result = 6 Then


                Else
                    Exit Sub
                End If


            
                Form1.query.executeQuery("DELETE FROM usr where id_usr=" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")
                        Me.ListaKontrahentow_Load(sender, e)

            End If
        Catch ex As Exception
            dodajedytujkon.Show()
        End Try
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)
        sortuj()

    End Sub
End Class