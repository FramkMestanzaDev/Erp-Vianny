Public Class vdetaleventa
    Dim ncom1 As String
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
    Public Property gncom1
        Get
            Return ncom1
        End Get
        Set(value)
            ncom1 = value
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
    Public Property glinea
        Get
            Return linea
        End Get
        Set(value)
            linea = value
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
    Public Property gventa_total
        Get
            Return venta_total
        End Get
        Set(value)
            venta_total = value
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
    Public Property gprecio_unitario2
        Get
            Return precio_unitario2
        End Get
        Set(value)
            precio_unitario2 = value
        End Set
    End Property

    Public Property gventa_total2
        Get
            Return venta_total2
        End Get
        Set(value)
            venta_total2 = venta_total2
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
    Sub New()

    End Sub
    Sub New(ByVal ncom1 As String, ByVal items As String, ByVal linea As String, ByVal sinonimo As String, ByVal cantidad As Double, ByVal precio_unitario As Double,
            ByVal venta_total As Double, ByVal valor_venta As Double, ByVal igv As Double, ByVal total As Double, ByVal precio_unitario2 As Double,
           ByVal venta_total2 As Double, ByVal valor_venta2 As Double, ByVal igv2 As Double, ByVal total2 As Double, ByVal op As String, ByVal almacen As String)

        gncom1 = ncom1
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
    End Sub
End Class
