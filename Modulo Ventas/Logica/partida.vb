Public Class partida
    Dim partida As String
    Dim nombre As String
    Dim codigo As String
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
        End Set
    End Property
    Public Property gnombre
        Get
            Return nombre
        End Get
        Set(value)
            nombre = value
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
    Sub New(ByVal partida As String, ByVal nombre As String, ByVal codigo As String)
        gpartida = partida
        gnombre = nombre
        gcodigo = codigo
    End Sub
End Class
