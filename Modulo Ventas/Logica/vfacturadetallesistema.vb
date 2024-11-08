Public Class vfacturadetallesistema
    Dim ncom As String
    Dim items As String
    Dim linea As String
    Dim sinonimo As String
    Dim cantidad As Double
    Dim precio_unitario As Double
    Dim venta_total As Double
    Dim valor_venta As Double
    Dim igv As Double
    Dim total As Double
    Dim precio_unitario2 As Double
    Dim venta_total2 As Double
    Dim valor_venta2 As Double
    Dim igv2 As Double
    Dim total2 As Double
    Dim op As String
    Dim almacen As String
    Dim rubro As String
    Dim um As String
    Dim grm As String
    Dim partida As String
    Dim fatura As String
    Dim correlativo_fac As String
    Dim tido As String
    Dim ano As String
    Dim ccia As String
    Dim rubrodetallado As String
    Public Property grubrodetallado
        Get
            Return rubrodetallado
        End Get
        Set(value)
            rubrodetallado = value
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
    Public Property gano
        Get
            Return ano
        End Get
        Set(value)
            ano = value
        End Set
    End Property
    Public Property gtido
        Get
            Return tido
        End Get
        Set(value)
            tido = value
        End Set
    End Property
    Public Property gfatura
        Get
            Return fatura
        End Get
        Set(value)
            fatura = value
        End Set
    End Property
    Public Property gcorrelativo_fac
        Get
            Return correlativo_fac
        End Get
        Set(value)
            correlativo_fac = value
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
    Public Property ggrm
        Get
            Return grm
        End Get
        Set(value)
            grm = value
        End Set
    End Property
    Public Property gum
        Get
            Return um
        End Get
        Set(value)
            um = value
        End Set
    End Property
    Public Property grubro
        Get
            Return rubro
        End Get
        Set(value)
            rubro = value
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
    Public Property gop
        Get
            Return op
        End Get
        Set(value)
            op = value
        End Set
    End Property
    Public Property gtotal2
        Get
            Return total2
        End Get
        Set(value)
            total2 = value
        End Set
    End Property
    Public Property gigv2
        Get
            Return igv2
        End Get
        Set(value)
            igv2 = value
        End Set
    End Property
    Public Property gvalor_venta2
        Get
            Return valor_venta2
        End Get
        Set(value)
            valor_venta2 = value
        End Set
    End Property
    Public Property gventa_total2
        Get
            Return venta_total2
        End Get
        Set(value)
            venta_total2 = value
        End Set
    End Property
    Public Property gprecio_unitario2
        Get
            Return precio_unitario2
        End Get
        Set(value)
            precio_unitario2 = value
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
    Public Property gigv
        Get
            Return igv
        End Get
        Set(value)
            igv = value
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

    Public Property gventa_total
        Get
            Return venta_total
        End Get
        Set(value)
            venta_total = value
        End Set
    End Property

    Public Property gcantidad
        Get
            Return cantidad
        End Get
        Set(value)
            cantidad = value
        End Set
    End Property

    Public Property gprecio_unitario
        Get
            Return precio_unitario
        End Get
        Set(value)
            precio_unitario = value
        End Set
    End Property

    Public Property gsinonimo
        Get
            Return sinonimo
        End Get
        Set(value)
            sinonimo = value
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
    Public Property gitems
        Get
            Return items
        End Get
        Set(value)
            items = value
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
    Sub New()

    End Sub
    Sub New(ByVal ncom As String,
    ByVal items As String,
        ByVal linea As String,
        ByVal sinonimo As String,
        ByVal cantidad As Double,
        ByVal precio_unitario As Double,
        ByVal venta_total As Double,
        ByVal valor_venta As Double,
        ByVal igv As Double,
        ByVal total As Double,
        ByVal precio_unitario2 As Double,
        ByVal venta_total2 As Double,
        ByVal valor_venta2 As Double,
        ByVal igv2 As Double,
        ByVal total2 As Double,
        ByVal op As String,
        ByVal almacen As String,
        ByVal rubro As String,
        ByVal um As String, ByVal ccia As String,
        ByVal grm As String, ByVal ano As String, ByVal rubrodetallado As String,
            ByVal partida As String, ByVal fatura As String, ByVal correlativo_fac As String, ByVal tido As String
   )
        grubrodetallado = rubrodetallado
        gccia = ccia
        gncom = ncom
        gitems = items
        glinea = linea
        gsinonimo = sinonimo
        gcantidad = cantidad
        gprecio_unitario = precio_unitario
        gventa_total = venta_total
        gvalor_venta = valor_venta
        gigv = igv
        gtotal = total
        gprecio_unitario2 = precio_unitario2
        gventa_total2 = venta_total2
        gvalor_venta2 = valor_venta2
        gigv2 = igv2
        gtotal2 = total2
        gop = op
        galmacen = almacen
        grubro = rubro
        gum = um
        ggrm = grm
        ggrm = partida
        gfatura = fatura
        gcorrelativo_fac = correlativo_fac
        gtido = tido
        gano = ano
    End Sub
End Class
