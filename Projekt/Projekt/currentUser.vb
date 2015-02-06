Public Class currentUser
    Private idusr As Long
    Private isadmin As Integer


    Public Sub setIDusr(ByRef id As Long)
        idusr = id
    End Sub

    Public Sub setIsAdmin(ByRef id As Integer)
        isadmin = id
    End Sub

    Function returnIdUser()
        Return idusr
    End Function

    Function returnIsAdmin()
        Return isadmin
    End Function


End Class