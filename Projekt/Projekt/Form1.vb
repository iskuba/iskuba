

Public Class Form1




    Public query As New Connection
    Public user As New currentUser




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        query.connectionString = _
    "Provider=Microsoft.Jet.OLEDB.4.0;User Id= zbud; Password=z; Data source=" & _
    "" & Application.StartupPath & "\baza\bazaERP.mdb; Jet OLEDB:System Database=" & Application.StartupPath & "\baza\Zabezpieczenia4.mdw;User ID=Paweł;Password=;"
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Trim(TextBox1.Text).Length = 0 Then
            MsgBox("Nie Podano Loginu!")
            Exit Sub
        End If

        If Trim(MaskedTextBox1.Text).Length = 0 Then
            MsgBox("Nie Podano Hasła!")
            Exit Sub
        End If


        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(MaskedTextBox1.Text)
        Dim haslo As String


        haslo = System.Convert.ToBase64String(byt)

        Dim wynik As Object

        wynik = query.wykonajZapytanie("SELECT idprc,admin from usr WHERE haslo ='" & haslo & "' and akronim='" & TextBox1.Text & "'")

        If wynik.GetType.FullName = GetType(DataTable).FullName Then
            Dim tabela As DataTable
            tabela = wynik

            If tabela.Rows.Count > 0 Then
                user.idusr = tabela.Rows(0).Item("idprc")
                user.isadmin = tabela.Rows(0).Item("admin")
                menuERP.ShowDialog()
            Else
                user = Nothing
                MsgBox("Podane dane są błędne!")
                Exit Sub
            End If
        End If

    End Sub
End Class
