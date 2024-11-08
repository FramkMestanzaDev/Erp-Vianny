Public Class vvventas
    Dim op As String
    Dim almacen As String
    Dim documento As String
    Dim serie As String
    Dim correla As String
    Dim codigo As String
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Public Property gop
        Get
            Return op
        End Get
        Set(value)
            op = value
        End Set
    End Property

    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
        End Set
    End Property

    Public Property gdocumento
        Get
            Return documento
        End Get
        Set(value)
            documento = value
        End Set
    End Property
    Public Property gserie
        Get
            Return serie
        End Get
        Set(value)
            serie = value
        End Set
    End Property
    Public Property gcorrela
        Get
            Return correla
        End Get
        Set(value)
            correla = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal op As String, ByVal almacen As String, ByVal documento As String, ByVal serie As String, ByVal correla As String, ByVal codigo As String)
        gop = op
        galmacen = almacen
        gdocumento = documento
        gserie = serie
        gcorrela = correla
        gcodigo = codigo
    End Sub
End Class
