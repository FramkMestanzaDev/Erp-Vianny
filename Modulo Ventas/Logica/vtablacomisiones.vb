Public Class vtablacomisiones
    Dim ccia As String
    Dim periodo As String
    Dim almacen As String
    Dim fecha As Date
    Dim rubro As String
    Dim id As String
    Dim fpagocontado As String
    Dim preciocontado As Double
    Dim operacioncontado As String
    Dim comisioncontado As Double
    Dim fpagocredito As String
    Dim preciocredito As Double
    Dim operacioncredito As String
    Dim comisioncredito As Double
    Dim rubro_codigo As String
    Public Property grubro_codigo
        Get
            Return rubro_codigo
        End Get
        Set(value)
            rubro_codigo = value
        End Set
    End Property
    Public Property gcomisioncredito
        Get
            Return comisioncredito
        End Get
        Set(value)
            comisioncredito = value
        End Set
    End Property
    Public Property goperacioncredito
        Get
            Return operacioncredito
        End Get
        Set(value)
            operacioncredito = value
        End Set
    End Property
    Public Property gpreciocredito
        Get
            Return preciocredito
        End Get
        Set(value)
            preciocredito = value
        End Set
    End Property
    Public Property gfpagocredito
        Get
            Return fpagocredito
        End Get
        Set(value)
            fpagocredito = value
        End Set
    End Property
    Public Property gcomisioncontado
        Get
            Return comisioncontado
        End Get
        Set(value)
            comisioncontado = value
        End Set
    End Property
    Public Property goperacioncontado
        Get
            Return operacioncontado
        End Get
        Set(value)
            operacioncontado = value
        End Set
    End Property
    Public Property gpreciocontado
        Get
            Return preciocontado
        End Get
        Set(value)
            preciocontado = value
        End Set
    End Property
    Public Property gfpagocontado
        Get
            Return fpagocontado
        End Get
        Set(value)
            fpagocontado = value
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
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
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
    Public Property gperiodo
        Get
            Return periodo
        End Get
        Set(value)
            periodo = value
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
    Public Property gid
        Get
            Return id
        End Get
        Set(value)
            id = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal periodo As String, ByVal ccia As String, ByVal fecha As Date, ByVal rubro As String, ByVal preciocontado As Double, ByVal preciocredito As Double, ByVal comisioncredito As Double, ByVal comisioncontado As Double,
             ByVal almacen As String, ByVal fpagocontado As String, ByVal operacioncontado As String, ByVal fpagocredito As String, ByVal id As String, ByVal rubro_codigo As String)
        gccia = ccia
        gperiodo = periodo
        galmacen = almacen
        gfecha = fecha
        grubro = rubro
        gfpagocontado = fpagocontado
        gpreciocontado = preciocontado
        goperacioncontado = operacioncontado
        gcomisioncontado = comisioncontado
        gfpagocredito = fpagocredito
        gpreciocredito = preciocredito
        goperacioncredito = operacioncredito
        gcomisioncredito = comisioncredito
        gid = id
        grubro_codigo = rubro_codigo
    End Sub
End Class
