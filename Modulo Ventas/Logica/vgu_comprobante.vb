Public Class vgu_comprobante
    Dim ncom As String
    Dim tdoc As String
    Dim fatura As String
    Dim correlativo_fac As String
    Dim fecha As Date
    Dim condicion_pago As String
    Dim ruc As String
    Dim razon_social As String
    Dim moneda As Integer
    Dim cambio As Double
    Dim venta_total As Double
    Dim venta_sinigv As Double
    Dim igv As Double
    Dim total_venta As Double
    Dim venta_total_do As Double
    Dim venta_sinigv_do As Double
    Dim igv_do As Double
    Dim total_venta_do As Double
    Dim glosa As String
    Dim op As String
    Dim vendedor As String
    Dim almacen As String
    Dim tipo_venta As String
    Public Property gvendedor
        Get
            Return vendedor
        End Get
        Set(value)
            vendedor = value
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
    Public Property gigv
        Get
            Return igv
        End Get
        Set(value)
            igv = value
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
    Public Property gventa_total_do
        Get
            Return venta_total_do
        End Get
        Set(value)
            venta_total_do = value
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
    Public Property gigv_do
        Get
            Return igv_do
        End Get
        Set(value)
            igv_do = value
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
    Public Property gglosa
        Get
            Return glosa
        End Get
        Set(value)
            glosa = value
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
    Public Property gtipo_venta
        Get
            Return tipo_venta
        End Get
        Set(value)
            tipo_venta = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal vendedor As String, ByVal ncom As String, ByVal tdoc As String, ByVal fatura As String, ByVal correlativo_fac As String, ByVal fecha As Date, ByVal condicion_pago As String, ByVal ruc As String, ByVal razon_social As String, ByVal moneda As Integer, ByVal cambio As Double, ByVal venta_total As Double, ByVal venta_sinigv As Double, ByVal igv As Double, ByVal total_venta As Double, ByVal venta_total_do As Double, ByVal venta_sinigv_do As Double, ByVal igv_do As Double, ByVal total_venta_do As Double, ByVal glosa As String, ByVal op As String, ByVal almacen As String, ByVal tipo_venta As String)
        gvendedor = vendedor
        gncom = ncom
        gtdoc = tdoc
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
        gop = op
        galmacen = almacen
        gtipo_venta = tipo_venta
    End Sub
End Class
