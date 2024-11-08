Public Class vfacturasistema
    Dim ncom As String
    Dim tdoc As String
    Dim fatura As String
    Dim correlativo_fac As String
    Dim fecha As Date
    Dim condicion_pago As String
    Dim ruc As String
    Dim razon_social As String
    Dim moneda As String
    Dim cambio As String
    Dim venta_total As Double
    Dim venta_sinigv As Double
    Dim igv As Double
    Dim total_venta As Double
    Dim venta_total_do As Double
    Dim venta_sinigv_do As Double
    Dim igv_do As Double
    Dim total_venta_do As Double
    Dim glosa As String
    Dim vendedor As String
    Dim almacen As String
    Dim tipo_venta As String
    Dim afecto As String
    Dim incluido As String
    Dim glosafac As String
    Dim fechafin As Date
    Dim codigo_sin As String
    Dim item_sin As String
    Dim nomb_sin As String
    Dim sguia As String
    Dim cguia As String
    Dim sfactura As String
    Dim cfactura As String
    Dim flag As String
    Dim ccia As String
    Dim stdoc As String
    Dim sfatura As String
    Dim scorrelativo_fac As String
    Dim oper As String
    Dim ano As String
    Dim pedido As String
    Dim rubro_detallado As String
    Public Property grubro_detallado
        Get
            Return rubro_detallado
        End Get
        Set(value)
            rubro_detallado = value
        End Set
    End Property
    Public Property gpedido
        Get
            Return pedido
        End Get
        Set(value)
            pedido = value
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
    Public Property goper
        Get
            Return oper
        End Get
        Set(value)
            oper = value
        End Set
    End Property
    Public Property gstdoc
        Get
            Return stdoc
        End Get
        Set(value)
            stdoc = value
        End Set
    End Property
    Public Property gsfatura
        Get
            Return sfatura
        End Get
        Set(value)
            sfatura = value
        End Set
    End Property
    Public Property gscorrelativo_fac
        Get
            Return scorrelativo_fac
        End Get
        Set(value)
            scorrelativo_fac = value
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
    Public Property gflag
        Get
            Return flag
        End Get
        Set(value)
            flag = value
        End Set
    End Property
    Public Property gsguia
        Get
            Return sguia
        End Get
        Set(value)
            sguia = value
        End Set
    End Property
    Public Property gcguia
        Get
            Return cguia
        End Get
        Set(value)
            cguia = value
        End Set
    End Property
    Public Property gsfactura
        Get
            Return sfactura
        End Get
        Set(value)
            sfactura = value
        End Set
    End Property
    Public Property gcfactura
        Get
            Return cfactura
        End Get
        Set(value)
            cfactura = value
        End Set
    End Property
    Public Property gcodigo_sin
        Get
            Return codigo_sin
        End Get
        Set(value)
            codigo_sin = value
        End Set
    End Property
    Public Property gitem_sin
        Get
            Return item_sin
        End Get
        Set(value)
            item_sin = value
        End Set
    End Property
    Public Property gnomb_sin
        Get
            Return nomb_sin
        End Get
        Set(value)
            nomb_sin = value
        End Set
    End Property
    Public Property gfechafin
        Get
            Return fechafin
        End Get
        Set(value)
            fechafin = value
        End Set
    End Property
    Public Property gtipo_venta
        Get
            Return tipo_venta
        End Get
        Set(value)
            tipo_venta = value
        End Set
    End Property
    Public Property gglosafac
        Get
            Return glosafac
        End Get
        Set(value)
            glosafac = value
        End Set
    End Property
    Public Property gincluido
        Get
            Return incluido
        End Get
        Set(value)
            incluido = value
        End Set
    End Property
    Public Property gafecto
        Get
            Return afecto
        End Get
        Set(value)
            afecto = value
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
    Public Property gvendedor
        Get
            Return vendedor
        End Get
        Set(value)
            vendedor = value
        End Set
    End Property
    Public Property gglosa
        Get
            Return glosa
        End Get
        Set(value)
            glosa = value
        End Set
    End Property
    Public Property gtotal_venta_do
        Get
            Return total_venta_do
        End Get
        Set(value)
            total_venta_do = value
        End Set
    End Property
    Public Property gigv_do
        Get
            Return igv_do
        End Get
        Set(value)
            igv_do = value
        End Set
    End Property
    Public Property gventa_sinigv_do
        Get
            Return venta_sinigv_do
        End Get
        Set(value)
            venta_sinigv_do = value
        End Set
    End Property
    Public Property gventa_total_do
        Get
            Return venta_total_do
        End Get
        Set(value)
            venta_total_do = value
        End Set
    End Property
    Public Property gtotal_venta
        Get
            Return total_venta
        End Get
        Set(value)
            total_venta = value
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
    Public Property gncom
        Get
            Return ncom
        End Get
        Set(value)
            ncom = value
        End Set
    End Property
    Public Property gtdoc
        Get
            Return tdoc
        End Get
        Set(value)
            tdoc = value
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
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Public Property gcondicion_pago
        Get
            Return condicion_pago
        End Get
        Set(value)
            condicion_pago = value
        End Set
    End Property
    Public Property gruc
        Get
            Return ruc
        End Get
        Set(value)
            ruc = value
        End Set
    End Property
    Public Property grazon_social
        Get
            Return razon_social
        End Get
        Set(value)
            razon_social = value
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
    Public Property gcambio
        Get
            Return cambio
        End Get
        Set(value)
            cambio = value
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
    Public Property gventa_sinigv
        Get
            Return venta_sinigv
        End Get
        Set(value)
            venta_sinigv = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ncom As String,
    ByVal tdoc As String,
        ByVal fatura As String,
        ByVal correlativo_fac As String,
        ByVal fecha As Date,
        ByVal condicion_pago As String,
        ByVal ruc As String,
        ByVal razon_social As String,
        ByVal moneda As String,
        ByVal cambio As String,
        ByVal venta_total As Double,
        ByVal venta_sinigv As Double,
        ByVal igv As Double,
        ByVal total_venta As Double,
        ByVal venta_total_do As Double,
        ByVal venta_sinigv_do As Double,
        ByVal igv_do As Double,
        ByVal total_venta_do As Double,
        ByVal glosa As String,
        ByVal vendedor As String,
        ByVal almacen As String,
        ByVal afecto As String,
        ByVal incluido As String,
        ByVal glosafac As String,
            ByVal tipo_venta As String,
            ByVal fechafin As String,
            ByVal codigo_sin As String,
            ByVal item_sin As String,
            ByVal nomb_sin As String,
          ByVal sguia As String,
ByVal cguia As String, ByVal pedido As String,
ByVal sfactura As String, ByVal rubro_detallado As String,
ByVal cfactura As String, ByVal ano As String,
            ByVal flag As String, ByVal ccia As String, ByVal stdoc As String, ByVal sfatura As String, ByVal scorrelativo_fac As String, ByVal oper As String
   )
        grubro_detallado = rubro_detallado
        gflag = flag
        gsguia = sguia
        gcguia = cguia
        gsfactura = sfactura
        gcfactura = cfactura
        gtdoc = tdoc
        gncom = ncom
        gfatura = fatura
        gcorrelativo_fac = correlativo_fac
        gfecha = fecha
        gcondicion_pago = condicion_pago
        gruc = ruc
        grazon_social = razon_social
        gmoneda = moneda
        gcambio = cambio
        gventa_total = venta_total
        gventa_sinigv = venta_sinigv
        gigv = igv
        gtotal_venta = total_venta
        gventa_total_do = venta_total_do
        gventa_sinigv_do = venta_sinigv_do
        gigv_do = igv_do
        gtotal_venta_do = total_venta_do
        gglosa = glosa
        gvendedor = vendedor
        galmacen = almacen
        gafecto = afecto
        gincluido = incluido
        gglosafac = glosafac
        gtipo_venta = tipo_venta
        gfechafin = fechafin
        gcodigo_sin = codigo_sin
        gitem_sin = item_sin
        gnomb_sin = nomb_sin
        gccia = ccia
        gstdoc = stdoc
        gsfatura = sfatura
        gscorrelativo_fac = scorrelativo_fac
        goper = oper
        gpedido = pedido
        gano = ano
    End Sub
End Class
