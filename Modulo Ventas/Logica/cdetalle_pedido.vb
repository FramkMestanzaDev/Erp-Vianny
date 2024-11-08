Public Class cdetalle_pedido
    Dim items As String
    Dim codigo_tela As String
    Dim descripcion As String
    Dim color As String
    Dim densidad As String
    Dim ancho As String
    Dim um As String
    Dim partida As String
    Dim almacen As String
    Dim cantidad As Double
    Dim precio_unitario As Double
    Dim igv As Double
    Dim Total As Double
    Dim subtotal As Double
    Dim numero_pedido As String
    Dim estado As String
    Dim estado2 As String
    Dim empresa As String
    Dim alm_num As String
    Dim rollos As String
    Public Property grollos
        Get
            Return rollos
        End Get
        Set(value)
            rollos = value
        End Set
    End Property
    Public Property galm_num
        Get
            Return alm_num
        End Get
        Set(value)
            alm_num = value
        End Set
    End Property
    Public Property gempresa
        Get
            Return empresa
        End Get
        Set(value)
            empresa = value
        End Set
    End Property
    Public Property gestado
        Get
            Return estado
        End Get
        Set(value)
            estado = value
        End Set
    End Property
    Public Property gestado2
        Get
            Return estado2
        End Get
        Set(value)
            estado2 = value
        End Set
    End Property
    Public Property gsubtotal
        Get
            Return subtotal
        End Get
        Set(value)
            subtotal = value
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
    Public Property gcodigo_tela
        Get
            Return codigo_tela
        End Get
        Set(value)
            codigo_tela = value
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
    Public Property gcolor
        Get
            Return color
        End Get
        Set(value)
            color = value
        End Set
    End Property
    Public Property gdensidad
        Get
            Return densidad
        End Get
        Set(value)
            densidad = value
        End Set
    End Property
    Public Property gancho
        Get
            Return ancho
        End Get
        Set(value)
            ancho = value
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
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
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
    Public Property gigv
        Get
            Return igv
        End Get
        Set(value)
            igv = value
        End Set
    End Property
    Public Property gTotal
        Get
            Return Total
        End Get
        Set(value)
            Total = value
        End Set
    End Property
    Public Property gnumero_pedido
        Get
            Return numero_pedido
        End Get
        Set(value)
            numero_pedido = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal items As String, ByVal codigo_tela As String, ByVal descripcion As String, ByVal color As String, ByVal densidad As String, ByVal ancho As String, ByVal um As String, ByVal partida As String, ByVal almacen As String, ByVal alm_num As String,
        ByVal cantidad As Double, ByVal precio_unitario As Double, ByVal igv As Double, ByVal rollos As String, ByVal total As Double, ByVal numero_pedido As String, ByVal subtotal As Double, ByVal estado As String, ByVal estado2 As String, ByVal empresa As String)
        gitems = items
        gcodigo_tela = codigo_tela
        gdescripcion = descripcion
        gcolor = color
        gdensidad = densidad
        gancho = ancho
        gum = um
        gpartida = partida
        galmacen = almacen
        gcantidad = cantidad
        gprecio_unitario = precio_unitario
        gigv = igv
        gTotal = total
        gnumero_pedido = numero_pedido
        gsubtotal = subtotal
        gestado = estado
        gestado2 = estado2
        gempresa = empresa
        galm_num = alm_num
        grollos = rollos
    End Sub
End Class
