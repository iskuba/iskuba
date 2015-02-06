''' <summary>
''' Klasa zawierająca inform
''' </summary>

Public Class currentUser

    ''' <summary>
    ''' zmienna przechowująca ID aktualnie zalogowanego uzytkownika
    ''' </summary>
    Private idusr As Long
    ''' <summary>
    ''' Gdy wartosc 0 brak uprawnien administratora
    ''' </summary>
    Private isadmin As Integer

    ''' <summary>
    ''' Metoda dopisująca do zmiennej id uzytkownika
    ''' </summary>
    Public Sub setIDusr(ByRef id As Long)
        idusr = id
    End Sub
    ''' <summary>
    ''' Metoda dopisująca do  isadmin  wartosc
    ''' </summary>
    Public Sub setIsAdmin(ByRef id As Integer)
        isadmin = id
    End Sub
    ''' <summary>
    ''' Metoda zwracająca id uzytkownika
    ''' </summary>
    Function returnIdUser()
        Return idusr
    End Function
    ''' <summary>
    ''' Metoda zwracająca wartosc isadmin
    ''' </summary>
    Function returnIsAdmin()
        Return isadmin
    End Function


End Class