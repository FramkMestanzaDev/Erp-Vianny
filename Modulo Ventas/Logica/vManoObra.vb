Public Class vManoObra


    Dim descripcion As String
    Dim merma As Double
    Dim consumo As Double
    Dim unidad As String
    Dim consumo_total As Double
    Dim costo_unitario As Double
    Dim costo_total As Double
    Dim id_cotizacion As String
    Public Property gcosto_total
        Get
            Return costo_total
        End Get
        Set(value)
            costo_total = value
        End Set
    End Property
    Public Property gcosto_unitario
        Get
            Return costo_unitario
        End Get
        Set(value)
            costo_unitario = value
        End Set
    End Property
    Public Property gmerma
        Get
            Return merma
        End Get
        Set(value)
            merma = value
        End Set
    End Property
    Public Property gconsumo
        Get
            Return consumo
        End Get
        Set(value)
            consumo = value
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
    Public Property gunidad
        Get
            Return unidad
        End Get
        Set(value)
            unidad = value
        End Set
    End Property
    Public Property gconsumo_total
        Get
            Return consumo_total
        End Get
        Set(value)
            consumo_total = value
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
    Sub New()

    End Sub
    Sub New(ByVal descripcion As String, ByVal merma As Double,
            ByVal consumo As Double, ByVal unidad As String, ByVal consumo_total As Double, ByVal costo_unitario As Double,
             ByVal id_cotizacion As String, ByVal costo_total As Double)

        gdescripcion = descripcion
        gmerma = merma
        gconsumo = consumo
        gunidad = unidad
        gconsumo_total = consumo_total
        gcosto_unitario = costo_unitario
        gid_cotizacion = id_cotizacion
        gcosto_total = costo_total
        gid_cotizacion = id_cotizacion
    End Sub
End Class
