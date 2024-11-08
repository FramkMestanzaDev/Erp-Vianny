Public Class vfacturavianny
    Dim mes As String
    Dim almacen As String
    Dim doc As String
    Dim serie As String
    Dim ano As String
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gano
        Get
            Return ano
        End Get
        Set(value)
            ano = value
        End Set
    End Property
    Public Property gdoc
        Get
            Return doc
        End Get
        Set(value)
            doc = value
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
    Public Property gmes
        Get
            Return mes
        End Get
        Set(value)
            mes = value
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
    Sub New()

    End Sub
    Sub New(ByVal mes As String, ByVal almacen As String, ByVal doc As String, ByVal serie As String, ByVal ano As String, ByRef ccia As String)
        gmes = mes
        galmacen = almacen
        gserie = serie
        gdoc = doc
        gano = ano
        gccia = ccia
    End Sub
End Class
