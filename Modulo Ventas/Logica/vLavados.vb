Public Class vLavados
    Dim codigo As String
    Dim descripcion As String
    Dim merma As Double
    Dim consumo As Double
    Dim unidad As String
    Dim consumo_tot As Double
    Dim costo_unitario As Double
    Dim costo_total As Double
    Dim items As String
    Dim id_cotizacion As String
    Public Property gid_cotizacion
        Get
            Return id_cotizacion
        End Get
        Set(value)
            id_cotizacion = value
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
    Public Property gitems
        Get
            Return items
        End Get
        Set(value)
            items = value
        End Set
    End Property
    Public Property gconsumo_tot
        Get
            Return consumo_tot
        End Get
        Set(value)
            consumo_tot = value
        End Set
    End Property
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
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
    Public Property gcosto_unitario
        Get
            Return costo_unitario
        End Get
        Set(value)
            costo_unitario = value
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
    Public Property gcosto_total
        Get
            Return costo_total
        End Get
        Set(value)
            costo_total = value
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
    Sub New()

    End Sub
    Sub New(ByVal codigo As String, ByVal descripcion As String, ByVal merma As Double, ByVal consumo As Double,
            ByVal unidad As String, ByVal consumo_tot As Double, ByVal costo_unitario As Double, ByVal costo_total As Double,
             ByVal items As String, ByVal id_cotizacion As String)
        gcodigo = codigo
        gdescripcion = descripcion
        gmerma = merma
        gconsumo = consumo
        gunidad = unidad
        gconsumo_tot = consumo_tot
        gcosto_unitario = costo_unitario
        gcosto_total = costo_total
        gitems = items
        gid_cotizacion = id_cotizacion
    End Sub
End Class
