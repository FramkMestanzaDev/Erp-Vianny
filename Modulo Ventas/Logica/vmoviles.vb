Public Class vmoviles
    Dim proveedor As String
    Dim linea As String
    Dim codigo As String
    Dim trabajador As String
    Dim area As String
    Dim mequipo As String
    Dim moequipo As String
    Dim preequipo As Double
    Dim cuotas As Integer
    Dim serie As String
    Dim plam As String
    Dim modalidad As String
    Dim empresa As String
    Dim fechas As Date
    Public Property gfechas
        Get
            Return fechas
        End Get
        Set(value)
            fechas = value
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
    Public Property gmodalidad
        Get
            Return modalidad
        End Get
        Set(value)
            modalidad = value
        End Set
    End Property
    Public Property gplam
        Get
            Return plam
        End Get
        Set(value)
            plam = value
        End Set
    End Property
    Public Property gserie
        Get
            Return serie
        End Get
        Set(value)
            serie = value
        End Set
    End Property
    Public Property gcuotas
        Get
            Return cuotas
        End Get
        Set(value)
            cuotas = value
        End Set
    End Property
    Public Property gpreequipo
        Get
            Return preequipo
        End Get
        Set(value)
            preequipo = value
        End Set
    End Property
    Public Property gmoequipo
        Get
            Return moequipo
        End Get
        Set(value)
            moequipo = value
        End Set
    End Property
    Public Property gmequipo
        Get
            Return mequipo
        End Get
        Set(value)
            mequipo = value
        End Set
    End Property
    Public Property garea
        Get
            Return area
        End Get
        Set(value)
            area = value
        End Set
    End Property
    Public Property gtrabajador
        Get
            Return trabajador
        End Get
        Set(value)
            trabajador = value
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
    Public Property gproveedor
        Get
            Return proveedor
        End Get
        Set(value)
            proveedor = value
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
    Sub New()

    End Sub
    Sub New(ByVal fini As Date,
        ByVal ffin As Date)
        gproveedor = proveedor
        glinea = linea
        gcodigo = codigo
        gtrabajador = trabajador
        garea = area
        gmequipo = mequipo
        gmoequipo = moequipo
        gpreequipo = preequipo
        gcuotas = cuotas
        gserie = serie
        gplam = plam
        gmodalidad = modalidad
        gempresa = empresa
        gfechas = fechas
    End Sub
End Class
