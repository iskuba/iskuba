Imports CrystalDecisions.CrystalReports.Engine


''' <summary>
''' Klasa przedstawiająca listę aktualnych towarow
''' </summary>

Public Class listatowarow
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


            wynik = oknoMenu.returnLogin.returnQuery.wykonajZapytanie("SELECT twrkarty.id as id,twrkod,twrnazwa,jmz,sum(ilosc) as iloscDostaw  FROM twrkarty LEFT JOIN traelem ON twrkarty.id = traelem.idtwr GROUP BY twrkarty.id,twrkod,twrnazwa,jmz;")
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

    ''' <summary>
    ''' Metoda sortująca datagrida
    ''' </summary>
    Sub sortuj()
        Dim table2 As DataTable
        table2 = DataGridView1.DataSource

        table2.DefaultView.RowFilter = "twrkod like '" & TextBox1.Text & "%' and twrnazwa like '" & TextBox2.Text & "%' "

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
    ''' Metoda tworzy obiekt dodajedytujtowar oraz wywołuje formę tej klasy
    ''' </summary>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim forma As New dodajedytujtowar
        forma.setOknoTowar(Me)
        forma.Show()
    End Sub
    ''' <summary>
    ''' Metoda edytuje wybrany z datagrida rekord oraz tworzy obiekt klasy dodajeydutjtowar i wywołuje formę tej klasy z id rekordu
    ''' </summary>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If IsNothing(DataGridView1.CurrentRow) Then

        Else
            On Error Resume Next
            Dim forma As New dodajedytujtowar
            forma.setOknoTowar(Me)
            forma.setId(DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value)
            forma.TextBox1.Text = DataGridView1.Item("twrkod", DataGridView1.CurrentRow.Index).Value
            forma.TextBox2.Text = DataGridView1.Item("twrnazwa", DataGridView1.CurrentRow.Index).Value
            forma.ComboBox1.Text = DataGridView1.Item("jmz", DataGridView1.CurrentRow.Index).Value
            forma.Show()

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

                result = MsgBox("Czy na pewno chcesz usunąć wybrany towar ?", vbYesNo)

                If result = 6 Then


                Else
                    Exit Sub
                End If


                Dim wynik As Object
                wynik = oknoMenu.returnLogin.returnQuery.wykonajZapytanie("SELECT * FROM  TraElem where idtwr =" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")

                If wynik.GetType.FullName = GetType(DataTable).FullName Then
                    Dim tabela As DataTable

                    tabela = wynik

                    If tabela.Rows.Count > 0 Then

                        MsgBox("Istnieją wiązania z wybranym towarem. Usuwanie rekordu zostanie anulowane !")
                    Else
                        oknoMenu.returnLogin.returnQuery.executeQuery("DELETE FROM twrkarty where id=" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")
                        Me.ListaKontrahentow_Load(sender, e)
                    End If


                End If

            End If
        Catch ex As Exception
            dodajedytujkon.Show()
        End Try
    End Sub
    ''' <summary>
    ''' sorotwanie on Textbox3textchange
    ''' </summary>
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)
        sortuj()

    End Sub
    ''' <summary>
    ''' Metoda generująca wydruk stanu magazynowego
    ''' </summary>
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim tabela As New DataTable
        tabela.Columns.Add("kod")
        tabela.Columns.Add("nazwa")
        tabela.Columns.Add("jmz")
        tabela.Columns.Add("ilosc")

        For i = 0 To DataGridView1.Rows.Count - 1

            tabela.Rows.Add(DataGridView1.Item("twrkod", i).Value, DataGridView1.Item("twrnazwa", i).Value, DataGridView1.Item("jmz", i).Value, DataGridView1.Item("iloscDostaw", i).Value)
        Next
        Dim rptDoc1 As ReportDocument = New ReportDocument()

        rptDoc1.Load(Application.StartupPath & "\raporty\stanMagazynu.rpt")
        rptDoc1.SetDataSource(tabela)

        Dim forma As New raport
        forma.CrystalReportViewer1.ReportSource = rptDoc1
        forma.Show()


    End Sub
End Class