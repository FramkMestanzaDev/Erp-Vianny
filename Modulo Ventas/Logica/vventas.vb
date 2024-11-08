Public Class vventas
    Dim fini As Date
    Dim ffin As Date
    Dim mes As String
    Dim VENDEDOR As String
    Dim op As String
    Dim almacen As String
    Dim documento As String
    Dim serie As String
    Dim correla As String
    Dim codigo As String
    Dim ANO As String
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gANO
        Get
            Return ANO
        End Get
        Set(value)
            ANO = value
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
    Public Property gVENDEDOR
        Get
            Return VENDEDOR
        End Get
        Set(value)
            VENDEDOR = value
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
    Public Property gfini
        Get
            Return fini
        End Get
        Set(value)
            fini = value
        End Set
    End Property
    Public Property gffin
        Get
            Return ffin
        End Get
        Set(value)
            ffin = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal fini As Date, ByVal ANO As String, ByVal ccia As String,
        ByVal ffin As Date, ByVal mes As String, ByVal VENDEDOR As String, ByVal op As String, ByVal almacen As String, ByVal documento As String, ByVal serie As String, ByVal correla As String, ByVal codigo As String)
        gfini = fini
        gffin = ffin
        gmes = mes
        gVENDEDOR = VENDEDOR
        gop = op
        galmacen = almacen
        gdocumento = documento
        gserie = serie
        gcorrela = correla
        gcodigo = codigo
        gANO = ANO
        gccia = ccia
    End Sub
End Class
