Public Class vtallas
    Dim codigo As String
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal codigo As String, ByVal gccia As String)
        gcodigo = codigo
        gccia = ccia
    End Sub
End Class
