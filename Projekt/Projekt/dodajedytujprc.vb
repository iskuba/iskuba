﻿''' <summary>
''' Klasa dodająca,edytująca pracownika
''' </summary>


Public Class dodajedytujprc

    ''' <summary>
    ''' Zmiena przechowująca referencję do listyPrc
    ''' </summary>
    Private oknoPrc As listapracownikow

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
    ''' metoda utawiająca zmieną oknoPrc
    ''' </summary>
    Public Sub setOknoTowar(ByRef state As listapracownikow)
        oknoPrc = state
    End Sub

    ''' <summary>
    ''' Metoda  wywoływująca edycję lub dodanie nowego rekordu , zależy czy id jest większe  czy równe 0 
    ''' </summary>

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(TextBox1.Text).Length = 0 Then
            MsgBox("Podaj Imię !")
            Exit Sub
        End If

        If Trim(TextBox8.Text).Length = 0 Then
            MsgBox("Podaj Nazwisko !")
            Exit Sub
        End If

        If Trim(TextBox2.Text).Length = 0 Or Trim(TextBox3.Text).Length = 0 Or Trim(TextBox4.Text).Length = 0 Or Trim(TextBox6.Text).Length = 0 Or Trim(TextBox7.Text).Length = 0 Then
            MsgBox("Podaj Dane Adresowe (Miejscowość, Kod Pocztowy , Numer Lokalu) !")
            Exit Sub
        End If



        If id = 0 Then '' jeżeli ID = 0 to oznacza ze tworzymy nowy rekord
            oknoPrc.returnOknoMenu.returnLogin.returnQuery.executeQuery("INSERT INTO prckarty (imie,miejscowosc,kodpocztowy,telefon,nrlokalu,ulica,nazwisko) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "-" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')")
            MsgBox("Dodano Rekord !")
            oknoPrc.ListaKontrahentow_Load(sender, e)
        Else
            oknoPrc.returnOknoMenu.returnLogin.returnQuery.executeQuery("UPDATE  prckarty set imie= '" & TextBox1.Text & "',miejscowosc='" & TextBox2.Text & "',kodpocztowy='" & TextBox3.Text & "-" & TextBox4.Text & "',telefon='" & TextBox5.Text & "',nrlokalu='" & TextBox6.Text & "',ulica='" & TextBox7.Text & "' ,nazwisko='" & TextBox8.Text & "' WHERE idprc=" & id & "")
            MsgBox("Rekord Został Zaktualizoany !")
            oknoPrc.ListaKontrahentow_Load(sender, e)
        End If


    End Sub
    ''' <summary>
    ''' Metoda sprawdzająca czy wprowadzona wartosc jest numeryczna
    ''' </summary>
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text.Length > 0 Then
            If IsNumeric(TextBox3.Text) Then
            Else
                TextBox3.Text = ""
            End If
        End If
    End Sub
    ''' <summary>
    ''' Metoda sprawdzająca czy wprowadzona wartosc jest numeryczna
    ''' </summary>
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text.Length > 0 Then
            If IsNumeric(TextBox4.Text) Then
            Else
                TextBox4.Text = ""
            End If
        End If
    End Sub
    ''' <summary>
    ''' Metoda ładująca formę naszej klasy
    ''' </summary>
    Private Sub dodajedytujkon_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If id = 0 Then
            Me.Text = " Dodaj Pracownika "
        Else
            Me.Text = " Edytuj Pracownika "
        End If
    End Sub
End Class