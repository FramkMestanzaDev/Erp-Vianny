Public Class vnia
    Dim mes As String
    Dim almacen As String
    Dim comprobante As String
    Dim almacen1 As String
    Dim ncom As String
    Dim guia As String
    Dim ano As String
    Dim partida As String
    Dim linea As String
    Dim ccom As String
    Dim ccia As String
    Dim op As String
    Public Property gop
        Get
            Return op
        End Get
        Set(value)
            op = value
        End Set
    End Property
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gccom
        Get
            Return ccom
        End Get
        Set(value)
            ccom = value
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
    Public Property glinea
        Get
            Return linea
        End Get
        Set(value)
            linea = value
        End Set

    End Property
    Public Property gguia
        Get
            Return guia
        End Get
        Set(value)
            guia = value
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
    Public Property gncom
        Get
            Return ncom
        End Get
        Set(value)
            ncom = value
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
    Public Property gcomprobante
        Get
            Return comprobante
        End Get
        Set(value)
            comprobante = value
        End Set
    End Property
    Public Property galmacen1
        Get
            Return almacen1
        End Get
        Set(value)
            almacen1 = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ruc As String, ByVal comprobante As String, ByVal almacen As String, ByVal partida As String, ByVal linea As String, ByVal ccia As String, ByVal op As String,
            ByVal almacen1 As String, ByVal ncom As String, ByVal guia As String, ByVal ano As String, ByVal ccom As String)
        gmes = ruc
        galmacen = almacen
        gcomprobante = comprobante
        galmacen1 = almacen1
        gncom = ncom
        gano = ano
        gguia = guia
        gpartida = partida
        glinea = linea
        gccom = ccom
        gccia = ccia
        gop = op
    End Sub
End Class
