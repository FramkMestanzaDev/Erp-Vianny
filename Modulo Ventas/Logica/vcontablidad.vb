Public Class vcontablidad
    Dim factura As String
    Dim periodo As String
    Dim registro As String
    Dim lacrado As String
    Dim lacrado_tesoreria
    Dim fecha As Date
    Dim MES As String
    Dim grm As String
    Dim ncom As String
    Dim almacen As String
    Dim doc As String
    Dim ruc As String

    Dim cia As String
    Dim tienda As String

    Dim serieb As String
    Dim serief As String

    Dim PARCIAL As Double
    Dim OBSERVACION As String
    Dim ID As Integer
    Dim COMPROBANTE As String
    Public Property gID
        Get
            Return ID
        End Get
        Set(value)
            ID = value
        End Set
    End Property
    Public Property gCOMPROBANTE
        Get
            Return COMPROBANTE
        End Get
        Set(value)
            COMPROBANTE = value
        End Set
    End Property
    Public Property gOBSERVACION
        Get
            Return OBSERVACION
        End Get
        Set(value)
            OBSERVACION = value
        End Set
    End Property
    Public Property gPARCIAL
        Get
            Return PARCIAL
        End Get
        Set(value)
            PARCIAL = value
        End Set
    End Property
    Public Property gserief
        Get
            Return serief
        End Get
        Set(value)
            serief = value
        End Set
    End Property
    Public Property gserieb
        Get
            Return serieb
        End Get
        Set(value)
            serieb = value
        End Set
    End Property
    Public Property gtienda
        Get
            Return tienda
        End Get
        Set(value)
            tienda = value
        End Set
    End Property
    Public Property gcia
        Get
            Return cia
        End Get
        Set(value)
            cia = value
        End Set
    End Property
    Public Property gdoc
        Get
            Return doc
        End Get
        Set(value)
            doc = value
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

    Public Property gMES
        Get
            Return MES
        End Get
        Set(value)
            MES = value
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
    Public Property glacrado_tesoreria
        Get
            Return lacrado_tesoreria
        End Get
        Set(value)
            lacrado_tesoreria = value
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
    Public Property gperiodo
        Get
            Return periodo
        End Get
        Set(value)
            periodo = value
        End Set
    End Property
    Public Property gregistro
        Get
            Return registro
        End Get
        Set(value)
            registro = value
        End Set
    End Property
    Public Property glacrado
        Get
            Return lacrado
        End Get
        Set(value)
            lacrado = value
        End Set
    End Property
    Public Property galmacen
        Get
            Return Almacen
        End Get
        Set(value)
            Almacen = value
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
    Public Property ggrm
        Get
            Return grm
        End Get
        Set(value)
            grm = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal factura As String, ByVal cotizacion As String, ByVal nia As String, ByVal dir_nia As String, ByVal lacrado_tesoreria As String, ByVal fecha As Date, ByVal MES As String,
            ByVal grm As String, ByVal almacen As String, ByVal ncom As String, ByVal doc As String, ByVal ruc As String, ByVal cia As String, ByVal tienda As String, ByVal serieb As String, ByVal serief As String,
            ByVal PARCIAL As Double, ByVal OBSERVACION As String, ByVal ID As Integer, ByVal COMPROBANTE As String)
        gfactura = factura
        gperiodo = periodo
        gregistro = registro
        glacrado = lacrado
        glacrado_tesoreria = lacrado_tesoreria
        gfecha = fecha
        gMES = MES
        ggrm = grm
        galmacen = almacen
        gncom = ncom
        gdoc = doc
        gruc = ruc
        gcia = cia
        gtienda = tienda
        gserieb = serieb
        gserief = serief
        gPARCIAL = PARCIAL
        gOBSERVACION = OBSERVACION
        gID = ID
        gCOMPROBANTE = COMPROBANTE
    End Sub
End Class
