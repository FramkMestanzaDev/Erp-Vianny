Public Class vventas_n
    Dim periodo As String
    Dim salida As String
    Dim comprobante As String
    Dim cliente As String
    Dim vendedor As String
    Dim fpago As String
    Dim valor_venta As Double
    Dim igv As Double
    Dim total As Double
    Dim pagado As Double
    Dim parcial As Double
    Dim observacion As String
    Dim fecha As Date
    Dim guia_nsa As String
    Dim fini As Date
    Dim ffin As Date
    Dim fichas As String
    Dim TIPO As String
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gTIPO
        Get
            Return TIPO
        End Get
        Set(value)
            TIPO = value
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
    Public Property gfichas
        Get
            Return fichas
        End Get
        Set(value)
            fichas = value
        End Set
    End Property
    Public Property gguia_nsa
        Get
            Return guia_nsa
        End Get
        Set(value)
            guia_nsa = value
        End Set
    End Property
    Public Property gperiodo
        Get
            Return periodo
        End Get
        Set(value)
            periodo = value
        End Set
    End Property
    Public Property gsalida
        Get
            Return salida
        End Get
        Set(value)
            salida = value
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
    Public Property gcliente
        Get
            Return cliente
        End Get
        Set(value)
            cliente = value
        End Set
    End Property
    Public Property gvendedor
        Get
            Return vendedor
        End Get
        Set(value)
            vendedor = value
        End Set
    End Property
    Public Property gfpago
        Get
            Return fpago
        End Get
        Set(value)
            fpago = value
        End Set
    End Property
    Public Property gvalor_venta
        Get
            Return valor_venta
        End Get
        Set(value)
            valor_venta = value
        End Set
    End Property
    Public Property gigv
        Get
            Return igv
        End Get
        Set(value)
            igv = value
        End Set
    End Property
    Public Property gtotal
        Get
            Return total
        End Get
        Set(value)
            total = value
        End Set
    End Property
    Public Property gpagado
        Get
            Return pagado
        End Get
        Set(value)
            pagado = value
        End Set
    End Property
    Public Property gparcial
        Get
            Return parcial
        End Get
        Set(value)
            parcial = value
        End Set
    End Property
    Public Property gobservacion
        Get
            Return observacion
        End Get
        Set(value)
            observacion = value
        End Set
    End Property
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal periodo As String, ByVal salida As String, ByVal comprobante As Double, ByVal guia_nsa As String, ByVal fichas As String, ByVal fini As Date, ByVal TIPO As String,
            ByVal cliente As Double, ByVal vendedor As String, ByVal fpago As String, ByVal valor_venta As Double, ByVal ffin As Date, ByVal ccia As String,
            ByVal igv As Double, ByVal total As Double, ByVal pagado As Double, ByVal parcial As Double, ByVal observacion As String, ByVal fecha As Date)
        gperiodo = periodo
        gsalida = salida
        gcomprobante = comprobante
        gcliente = cliente
        gvendedor = vendedor
        gfpago = fpago
        gvalor_venta = valor_venta
        gigv = igv
        gtotal = total
        gpagado = pagado
        gparcial = parcial
        gobservacion = observacion
        gfecha = fecha
        gguia_nsa = guia_nsa
        gfini = fini
        gfini = fini
        gfichas = fichas
        gTIPO = TIPO
        gccia = ccia
    End Sub
End Class
