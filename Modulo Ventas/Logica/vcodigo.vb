Public Class vcodigo
    Dim ccia As String
    Dim linea As String
    Dim linea2 As String
    Dim familia As String
    Dim subfamilia As String
    Dim codcolor As String
    Dim color As String
    Dim ancho As String
    Dim densidad As Double
    Dim detalle As String
    Dim descripcion As String
    Dim nombre_comercial As String
    Dim um As String
    Dim TipCtrl_17 As String
    Dim almacen As String
    Dim tipo As String
    Dim barra As String
    Dim rubro As String
    Public Property grubro
        Get
            Return rubro
        End Get
        Set(value)
            rubro = value
        End Set
    End Property
    Public Property gbarra
        Get
            Return barra
        End Get
        Set(value)
            barra = value
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
    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
        End Set
    End Property
    Public Property gTipCtrl_17
        Get
            Return TipCtrl_17
        End Get
        Set(value)
            TipCtrl_17 = value
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
    Public Property gnombre_comercial
        Get
            Return nombre_comercial
        End Get
        Set(value)
            nombre_comercial = value
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
    Public Property gdetalle
        Get
            Return detalle
        End Get
        Set(value)
            detalle = value
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
    Public Property gcolor
        Get
            Return color
        End Get
        Set(value)
            color = value
        End Set
    End Property
    Public Property gcodcolor
        Get
            Return codcolor
        End Get
        Set(value)
            codcolor = value
        End Set
    End Property
    Public Property gsubfamilia
        Get
            Return subfamilia
        End Get
        Set(value)
            subfamilia = value
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
    Public Property glinea
        Get
            Return linea
        End Get
        Set(value)
            linea = value
        End Set
    End Property
    Public Property glinea2
        Get
            Return linea2
        End Get
        Set(value)
            linea2 = value
        End Set
    End Property
    Public Property gfamilia
        Get
            Return familia
        End Get
        Set(value)
            familia = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ccia As String, ByVal linea As String, ByVal familia As String, ByVal subfamilia As String, ByVal codcolor As String, ByVal color As String, ByVal ancho As String, ByVal densidad As Double, ByVal detalle As String, ByVal rubro As String,
            ByVal descripcion As String, ByVal nombre_comercial As String, ByVal um As String, ByVal TipCtrl_17 As String, ByVal almacen As String, ByVal tipo As String, ByVal linea2 As String, ByVal barra As String)
        gccia = ccia
        glinea = linea
        gfamilia = familia
        gsubfamilia = subfamilia
        gcodcolor = codcolor
        gcolor = color
        gancho = ancho
        gdensidad = densidad
        gdetalle = detalle
        gdescripcion = descripcion
        gnombre_comercial = nombre_comercial
        gum = um
        gTipCtrl_17 = TipCtrl_17
        galmacen = almacen
        gtipo = tipo
        glinea2 = linea2
        gbarra = barra
        grubro = rubro
    End Sub
End Class
