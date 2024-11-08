Public Class vpackin_tela_cab
    Dim id_packing As String
    Dim cod_pedido As Integer
    Dim nota_pedido As String
    Dim codigo_pro As String
    Dim descrip_pro As String
    Dim cantidad  As Double
    Dim f_despacho As Date
    Dim f_actual As Date
    Dim cliente As String
    Dim Vendedor  As String
    Dim partida As String
    Dim kg_seleccionados As Double
    Dim estado As String
    Dim almacen As String
    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
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
    Public Property gcliente
        Get
            Return cliente
        End Get
        Set(value)
            cliente = value
        End Set
    End Property
    Public Property gid_packing
        Get
            Return id_packing
        End Get
        Set(value)
            id_packing = value
        End Set
    End Property
    Public Property gcod_pedido
        Get
            Return cod_pedido
        End Get
        Set(value)
            cod_pedido = value
        End Set
    End Property
    Public Property gnota_pedido
        Get
            Return nota_pedido
        End Get
        Set(value)
            nota_pedido = value
        End Set
    End Property
    Public Property gcodigo_pro
        Get
            Return codigo_pro
        End Get
        Set(value)
            codigo_pro = value
        End Set
    End Property
    Public Property gdescrip_pro
        Get
            Return descrip_pro
        End Get
        Set(value)
            descrip_pro = value
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
    Public Property gf_despacho
        Get
            Return f_despacho
        End Get
        Set(value)
            f_despacho = value
        End Set
    End Property
    Public Property gf_actual
        Get
            Return f_actual
        End Get
        Set(value)
            f_actual = value
        End Set
    End Property
    Public Property gVendedor
        Get
            Return Vendedor
        End Get
        Set(value)
            Vendedor = value
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
    Public Property gkg_seleccionados
        Get
            Return kg_seleccionados
        End Get
        Set(value)
            kg_seleccionados = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal id_packing As String, ByVal cod_pedido As Integer, ByVal nota_pedido As String, ByVal estado As String,
            ByVal codigo_pro As String, ByVal descrip_pro As String, ByVal cantidad As Double,
            ByVal f_despacho As Date, ByVal f_actual As Date, ByVal cliente As String, ByVal Vendedor As String,
            ByVal partida As String, ByVal kg_seleccionados As Double, ByVal almacen As String)

        gid_packing = id_packing
        gcod_pedido = cod_pedido
        gnota_pedido = nota_pedido
        gcodigo_pro = codigo_pro
        gdescrip_pro = descrip_pro
        gcantidad = cantidad
        gf_despacho = f_despacho
        gf_actual = f_actual
        gcliente = cliente
        gVendedor = Vendedor
        gpartida = partida
        gkg_seleccionados = kg_seleccionados
        gestado = estado
        galmacen = almacen
    End Sub
End Class
