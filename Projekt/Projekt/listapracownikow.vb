''' <summary>
''' Klasa przedstawiająca listę aktualnych pracownikow
''' </summary>


Public Class listapracownikow

    ''' <summary>
    ''' Zmienna przechowująca referencję do menuERP
    ''' </summary>
    Private oknoMenu As menuERP
    ''' <summary>
    ''' Metoda ustawiająca zmieną oknoMenu
    ''' </summary>
    Public Sub setOknoMenu(ByRef state As menuERP)
        oknoMenu = state
    End Sub
    ''' <summary>
    ''' Metoda zwracająca zmieną oknoMenu
    ''' </summary>
    Function returnOknoMenu() As menuERP
        Return oknoMenu
    End Function


    ''' <summary>
    ''' Wywołanie metody ładującej formę klasy
    ''' </summary>
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
            wynik = oknoMenu.returnLogin.returnQuery.wykonajZapytanie("SELECT * FROM prckarty")
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


    ''' <summary>
    ''' Metoda sortująca datagrida
    ''' </summary>
    Sub sortuj()
        Dim table2 As DataTable
        table2 = DataGridView1.DataSource

        '' MsgBox("Kod_Towaru " & kod & " and Numer_Zlecenia " & zw & "")
        ''    MsgBox("Twr_Kod like '" & TextBox1.Text & "%' and NumerZlecenia like '" & TextBox2.Text & "%'and Dostawa like '" & TextBox3.Text & "%' and NumerPrzewodnika like '" & TextBox4.Text & "%' and  " & rw & "")

        table2.DefaultView.RowFilter = "imie like '" & TextBox1.Text & "%' and nazwisko like '" & TextBox2.Text & "%' and miejscowosc like '" & TextBox3.Text & "%' "

    End Sub

    ''' <summary>
    ''' sortowanie on TextBox1TextChange
    ''' </summary>
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        sortuj()
    End Sub
    ''' <summary>
    ''' sortowanie on TextBox2TextChange
    ''' </summary>
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        sortuj()
    End Sub
    ''' <summary>
    ''' Metoda tworzy obiekt dodajedytujprc oraz wywołuje formę tej klasy
    ''' </summary>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim forama As New dodajedytujprc
        forama.setOknoTowar(Me)
        forama.Show()

    End Sub


    ''' <summary>
    ''' Metoda edytuje wybrany z datagrida rekord oraz tworzy obiekt klasy dodajeydutjdostawe i wywołuje formę tej klasy z id rekordu
    ''' </summary>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If IsNothing(DataGridView1.CurrentRow) Then

        Else
            On Error Resume Next
            Dim forama As New dodajedytujprc
            forama.setOknoTowar(Me)



            forama.setId(DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value)
            forama.TextBox1.Text = DataGridView1.Item("imie", DataGridView1.CurrentRow.Index).Value
            forama.TextBox2.Text = DataGridView1.Item("miejscowosc", DataGridView1.CurrentRow.Index).Value
            Dim kod As String

            kod = DataGridView1.Item("kodpocztowy", DataGridView1.CurrentRow.Index).Value
            forama.TextBox3.Text = kod(0)
            forama.TextBox4.Text = kod(1)
            forama.TextBox5.Text = DataGridView1.Item("telefon", DataGridView1.CurrentRow.Index).Value
            forama.TextBox6.Text = DataGridView1.Item("nrlokalu", DataGridView1.CurrentRow.Index).Value
            forama.TextBox7.Text = DataGridView1.Item("ulica", DataGridView1.CurrentRow.Index).Value
            forama.TextBox8.Text = DataGridView1.Item("nazwisko", DataGridView1.CurrentRow.Index).Value
            forama.Show()

        End If

    End Sub
    ''' <summary>
    ''' Metoda usuwająca aktualnie wybrany element z datagirida
    ''' </summary>
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
                wynik = oknoMenu.returnLogin.returnQuery.wykonajZapytanie("SELECT idprc FROM  usr where idprc =" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & " UNION ALL SELECT idpracownika FROM TraNag Where idpracownika =" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")

                If wynik.GetType.FullName = GetType(DataTable).FullName Then
                    Dim tabela As DataTable

                    tabela = wynik

                    If tabela.Rows.Count > 0 Then

                        MsgBox("Istnieją wiązania z wybranym pracownikiem. Usuwanie rekordu zostanie anulowane !")
                    Else

                        oknoMenu.returnLogin.returnQuery.executeQuery("DELETE FROM prckarty where idprc=" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")
                        Me.ListaKontrahentow_Load(sender, e)
                    End If


                End If

            End If
        Catch ex As Exception

        End Try
    End Sub
    ''' <summary>
    ''' sorotwanie on Textbox3textchange
    ''' </summary>
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        sortuj()

    End Sub
End Class