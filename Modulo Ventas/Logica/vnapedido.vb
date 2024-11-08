Public Class vnapedido

    Dim numero_pedido As String
    Dim fecha As Date
    Dim codigo_cli As String
    Dim moneda As String
    Dim responsable As String
    Dim valor_comercial As String
    Dim dir_despacho As String
    Dim forma_pago As String
    Dim vendedor As String
    Dim or_compra As String
    Dim f_enrtrega As Date
    Dim inmediato As String
    Dim observaciones As String
    Dim sub_total As Double
    Dim igv As Double
    Dim total As Double
    Dim estado As String
    Dim preinc As String
    Dim escom As String
    Dim esalm As String
    Dim esttrans As String
    Dim almacen As String
    Dim cotizacion As String
    Dim factura As String
    Dim periodo As String
    Dim empresa As String
    Dim Rollos As String
    Public Property gRollos
        Get
            Return Rollos
        End Get
        Set(value)
            Rollos = value
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
    Public Property gperiodo
        Get
            Return periodo
        End Get
        Set(value)
            periodo = value
        End Set
    End Property
    Public Property gcotizacion
        Get
            Return cotizacion
        End Get
        Set(value)
            cotizacion = value
        End Set
    End Property
    Public Property gfactura
        Get
            Return factura
        End Get
        Set(value)
            factura = value
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
    Public Property gescom
        Get
            Return escom
        End Get
        Set(value)
            escom = value
        End Set
    End Property
    Public Property gesalm
        Get
            Return esalm
        End Get
        Set(value)
            esalm = value
        End Set
    End Property
    Public Property gesttrans
        Get
            Return esttrans
        End Get
        Set(value)
            esttrans = value
        End Set
    End Property
    Public Property gpreinc
        Get
            Return preinc
        End Get
        Set(value)
            preinc = value
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
    Public Property gobservaciones
        Get
            Return observaciones
        End Get
        Set(value)
            observaciones = value
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
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Public Property gcodigo_cli
        Get
            Return codigo_cli
        End Get
        Set(value)
            codigo_cli = value
        End Set
    End Property
    Public Property gmoneda
        Get
            Return moneda
        End Get
        Set(value)
            moneda = value
        End Set
    End Property
    Public Property gresponsable
        Get
            Return responsable
        End Get
        Set(value)
            responsable = value
        End Set
    End Property
    Public Property gvalor_comercial
        Get
            Return valor_comercial
        End Get
        Set(value)
            valor_comercial = value
        End Set
    End Property
    Public Property gdir_despacho
        Get
            Return dir_despacho
        End Get
        Set(value)
            dir_despacho = value
        End Set
    End Property
    Public Property gforma_pago
        Get
            Return forma_pago
        End Get
        Set(value)
            forma_pago = value
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
    Public Property gor_compra
        Get
            Return or_compra
        End Get
        Set(value)
            or_compra = value
        End Set
    End Property
    Public Property gf_enrtrega
        Get
            Return f_enrtrega
        End Get
        Set(value)
            f_enrtrega = value
        End Set
    End Property
    Public Property ginmediato
        Get
            Return inmediato
        End Get
        Set(value)
            inmediato = value
        End Set
    End Property
    Public Property gsub_total
        Get
            Return sub_total
        End Get
        Set(value)
            sub_total = value
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
    Sub New()

    End Sub
    Sub New(ByVal numero_pedido As String, ByVal fecha As Date, ByVal codigo_cli As String, ByVal moneda As String, ByVal responsable As String, ByVal valor_comercial As String, ByVal dir_despacho As String, ByVal forma_pago As String, ByVal vendedor As String,
        ByVal or_compra As String, ByVal f_enrtrega As Date, ByVal inmediato As String, ByVal observaciones As String, ByVal sub_total As Double, ByVal igv As Double, ByVal total As Double, ByVal estado As String, ByVal preinc As String, ByVal escom As String, ByVal esalm As String, ByVal esttrans As String,
            ByVal almacen As String, ByVal cotizacion As String, ByVal factura As String, ByVal periodo As String, ByVal empresa As String, ByVal Rollos As String)
        gnumero_pedido = numero_pedido
        gfecha = fecha
        gcodigo_cli = codigo_cli
        gmoneda = moneda
        gresponsable = responsable
        gvalor_comercial = valor_comercial
        gdir_despacho = dir_despacho
        gforma_pago = forma_pago
        gvendedor = vendedor
        gor_compra = or_compra
        gf_enrtrega = f_enrtrega
        ginmediato = inmediato
        gobservaciones = observaciones
        gsub_total = sub_total
        gigv = igv
        gtotal = total
        gestado = estado
        gpreinc = preinc
        gescom = escom
        gesalm = esalm
        gesttrans = esttrans
        galmacen = almacen
        gfactura = factura
        gcotizacion = cotizacion
        gperiodo = periodo
        gempresa = empresa
        gRollos = Rollos
    End Sub
End Class
