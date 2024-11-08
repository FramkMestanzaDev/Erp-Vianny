Public Class vproceso
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ccia As String)

        gccia = ccia

    End Sub
End Class
