Public Class fadat
    Dim partida As String
    Dim codigo As String
    Dim descripcion As String
    Dim rollo As String
    Dim empresa As String
    Public Property gempresa
        Get
            Return empresa
        End Get
        Set(value)
            empresa = value
        End Set
    End Property
    Public Property grollo
        Get
            Return rollo
        End Get
        Set(value)
            rollo = value
        End Set
    End Property
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
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
    Public Property gdescripcion
        Get
            Return descripcion
        End Get
        Set(value)
            descripcion = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal codigo As String, ByVal descripcion As String, ByVal rollo As String, ByVal empresa As String)
        gpartida = partida
        gcodigo = codigo
        gdescripcion = descripcion
        grollo = rollo
        gempresa = empresa
    End Sub
End Class
