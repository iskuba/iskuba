''' <summary>
''' Klasa dodająca,edytująca towar
''' </summary>


Public Class dodajedytujtowar

    ''' <summary>
    ''' Zmiena przechowująca referencję do listyTowarow
    ''' </summary>
    Private oknoTowar As listatowarow

    ''' <summary>
    ''' zmienna przechowująca id wybranego rekordu, jeżeli id > 0 oznacza to iż rekord jest edytowany 
    ''' </summary>
    Private id As Long

    ''' <summary>
    ''' metoda ustawiająca id rekordu
    ''' </summary>
    Public Sub setId(ByVal state As Long)
        id = state
    End Sub
    ''' <summary>
    ''' metoda utawiająca zmieną oknoTowar
    ''' </summary>
    Public Sub setOknoTowar(ByRef state As listatowarow)
        oknoTowar = state
    End Sub


    ''' <summary>
    ''' Metoda  wywoływująca edycję lub dodanie nowego rekordu , zależy czy id jest większe  czy równe 0 
    ''' </summary>
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


    ''' <summary>
    ''' Metoda ładująca formę naszej klasy
    ''' </summary>
    Private Sub dodajedytujkon_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        If id = 0 Then
            Me.Text = " Dodaj Towar"
        Else
            Me.Text = " Edytuj Towar"
        End If
    End Sub
End Class