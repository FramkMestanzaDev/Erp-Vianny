Public Class VCOTIZACION
    Dim id_cotizacion As String
    Dim cliente As String
    Dim descripcion As String
    Dim estilo As String
    Dim rango_tallas As String
    Dim ti_cambio As Double
    Dim moneda As String
    Dim linea As String
    Dim gasto_logistico As Double
    Dim gasto_operativo As Double
    Dim gasto_administrativo As Double
    Dim gasto_financiero As Double
    Dim gasto_venta As Double
    Dim costo_producto As Double
    Dim margen_utilidad As Double
    Dim valor_venta As Double
    Dim igv As Double
    Dim precio_venta As Double
    Dim estado As String
    Dim imagen As String
    Dim costot_tela As Double
    Dim costot_AviosC As Double
    Dim costot_AviosA As Double
    Dim costot_Lavado As Double
    Dim costot_Aplicaciones As Double
    Dim costot_MObra As Double
    Dim vendedor As String
    Dim fecha As Date
    Dim finicial As Date
    Dim ffinal As Date
    Dim tipo As String
    Dim margen_utilidad_moneda As Double
    Dim od As String
    Dim od_version As String
    Dim Pm As String
    Dim Tela As String
    Dim empresa As String
    Dim proveedor As String
    Dim articulo As String
    Dim flimite As Date
    Public Property gproveedor
        Get
            Return proveedor
        End Get
        Set(value)
            proveedor = value
        End Set
    End Property
    Public Property garticulo
        Get
            Return articulo
        End Get
        Set(value)
            articulo = value
        End Set
    End Property
    Public Property gflimite
        Get
            Return flimite
        End Get
        Set(value)
            flimite = value
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
    Public Property god
        Get
            Return od
        End Get
        Set(value)
            od = value
        End Set
    End Property
    Public Property god_version
        Get
            Return od_version
        End Get
        Set(value)
            od_version = value
        End Set
    End Property
    Public Property gPm
        Get
            Return Pm
        End Get
        Set(value)
            Pm = value
        End Set
    End Property
    Public Property gTela
        Get
            Return Tela
        End Get
        Set(value)
            Tela = value
        End Set
    End Property
    Public Property gmargen_utilidad_moneda
        Get
            Return margen_utilidad_moneda
        End Get
        Set(value)
            margen_utilidad_moneda = value
        End Set
    End Property
    Public Property ggasto_financiero
        Get
            Return gasto_financiero
        End Get
        Set(value)
            gasto_financiero = value
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
    Public Property gtipo
        Get
            Return tipo
        End Get
        Set(value)
            tipo = value
        End Set
    End Property
    Public Property gfinicial
        Get
            Return finicial
        End Get
        Set(value)
            finicial = value
        End Set
    End Property
    Public Property gffinal
        Get
            Return ffinal
        End Get
        Set(value)
            ffinal = value
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
    Public Property gvendedor
        Get
            Return vendedor
        End Get
        Set(value)
            vendedor = value
        End Set
    End Property
    Public Property gid_cotizacion
        Get
            Return id_cotizacion
        End Get
        Set(value)
            id_cotizacion = value
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
    Public Property gdescripcion
        Get
            Return descripcion
        End Get
        Set(value)
            descripcion = value
        End Set
    End Property
    Public Property gestilo
        Get
            Return estilo
        End Get
        Set(value)
            estilo = value
        End Set
    End Property
    Public Property grango_tallas
        Get
            Return rango_tallas
        End Get
        Set(value)
            rango_tallas = value
        End Set
    End Property
    Public Property gti_cambio
        Get
            Return ti_cambio
        End Get
        Set(value)
            ti_cambio = value
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
    Public Property ggasto_venta
        Get
            Return gasto_venta
        End Get
        Set(value)
            gasto_venta = value
        End Set
    End Property
    Public Property ggasto_operativo
        Get
            Return gasto_operativo
        End Get
        Set(value)
            gasto_operativo = value
        End Set
    End Property
    Public Property ggasto_logistico
        Get
            Return gasto_logistico
        End Get
        Set(value)
            gasto_logistico = value
        End Set
    End Property
    Public Property ggasto_administrativo
        Get
            Return gasto_administrativo
        End Get
        Set(value)
            gasto_administrativo = value
        End Set
    End Property
    Public Property gcosto_producto
        Get
            Return costo_producto
        End Get
        Set(value)
            costo_producto = value
        End Set
    End Property
    Public Property gmargen_utilidad
        Get
            Return margen_utilidad
        End Get
        Set(value)
            margen_utilidad = value
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
    Public Property gprecio_venta
        Get
            Return precio_venta
        End Get
        Set(value)
            precio_venta = value
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
    Public Property gimagen
        Get
            Return imagen
        End Get
        Set(value)
            imagen = value
        End Set
    End Property
    Public Property gcostot_tela
        Get
            Return costot_tela
        End Get
        Set(value)
            costot_tela = value
        End Set
    End Property
    Public Property gcostot_AviosC
        Get
            Return costot_AviosC
        End Get
        Set(value)
            costot_AviosC = value
        End Set
    End Property
    Public Property gcostot_AviosA
        Get
            Return costot_AviosA
        End Get
        Set(value)
            costot_AviosA = value
        End Set
    End Property
    Public Property gcostot_Lavado
        Get
            Return costot_Lavado
        End Get
        Set(value)
            costot_Lavado = value
        End Set
    End Property
    Public Property gcostot_Aplicaciones
        Get
            Return costot_Aplicaciones
        End Get
        Set(value)
            costot_Aplicaciones = value
        End Set
    End Property
    Public Property gcostot_MObra
        Get
            Return costot_MObra
        End Get
        Set(value)
            costot_MObra = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal id_cotizacion As String, ByVal cliente As String, ByVal descripcion As String, ByVal estilo As String, ByVal od As String, ByVal od_version As String,
            ByVal rango_tallas As String, ByVal ti_cambio As Double, ByVal moneda As String, ByVal gasto_logistico As Double, ByVal Pm As String,
            ByVal gasto_operativo As Double, ByVal gasto_administrativo As Double, ByVal gasto_financiero As Double, ByVal gasto_venta As Double, ByVal Tela As String,
            ByVal costo_producto As Double, ByVal margen_utilidad As Double, ByVal valor_venta As Double, ByVal igv As Double,
            ByVal precio_venta As Double, ByVal estado As String, ByVal imagen As String, ByVal costot_tela As Double, proveedor As String,
            ByVal costot_AviosC As Double, ByVal costot_AviosA As Double, ByVal costot_Lavado As Double, ByVal costot_Aplicaciones As Double, empresa As String,
          ByVal costot_MObra As Double, ByVal vendedor As String, ByVal fecha As Date, ByVal finicial As Date, ByVal flimite As Date, articulo As String,
           ByVal ffinal As Date, ByVal tipo As String, ByVal linea As String, ByVal margen_utilidad_moneda As Double)
        god = od
        god_version = od_version
        gPm = Pm
        gTela = Tela
        gid_cotizacion = id_cotizacion
        gcliente = cliente
        gdescripcion = descripcion
        gestilo = estilo
        grango_tallas = rango_tallas
        gti_cambio = ti_cambio
        gmoneda = moneda
        glinea = linea
        ggasto_logistico = gasto_logistico
        ggasto_operativo = gasto_operativo
        ggasto_administrativo = gasto_administrativo
        ggasto_financiero = gasto_financiero
        ggasto_venta = gasto_venta
        gcosto_producto = costo_producto
        gmargen_utilidad = margen_utilidad
        gvalor_venta = valor_venta
        gigv = igv
        gprecio_venta = precio_venta
        gestado = estado
        gimagen = imagen
        gcostot_tela = costot_tela
        gcostot_AviosC = costot_AviosC
        gcostot_AviosA = costot_AviosA
        gcostot_Lavado = costot_Lavado
        gcostot_Aplicaciones = costot_Aplicaciones
        gcostot_MObra = costot_MObra
        gvendedor = vendedor
        gfecha = fecha
        gfinicial = finicial
        gffinal = ffinal
        gtipo = tipo
        gmargen_utilidad_moneda = margen_utilidad_moneda
        gempresa = empresa
        gflimite = flimite
        garticulo = articulo
        gproveedor = proveedor
    End Sub
End Class
