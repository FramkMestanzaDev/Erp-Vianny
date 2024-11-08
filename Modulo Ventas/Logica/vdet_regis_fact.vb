Public Class vdet_regis_fact
    Dim periodo As String
    Dim comprobante As String
    Dim cliente As String
    Dim valor_venta As Double
    Dim igv As Double
    Dim total As Double
    Dim pagado As Double
    Dim parcial As Double
    Dim observacion As String
    Dim fecha As Date
    Public Property gperiodo
        Get
            Return periodo
        End Get
        Set(value)
            periodo = value
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
    Sub New(ByVal periodo As String, ByVal comprobante As String, ByVal cliente As String, ByVal valor_venta As Double,
             ByVal igv As Double, ByVal total As Double, ByVal pagado As Double, ByVal parcial As Double,
             ByVal observacion As String, ByVal fecha As Date)
        gperiodo = periodo
        gcomprobante = comprobante
        gcliente = cliente
        gvalor_venta = valor_venta
        gigv = igv
        gtotal = total
        gpagado = pagado
        gparcial = parcial
        gobservacion = observacion
        gfecha = fecha
    End Sub
End Class
