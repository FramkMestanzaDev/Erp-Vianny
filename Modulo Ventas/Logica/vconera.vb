Public Class vconera
    Dim id_conera As String
    Dim juego_urdido As String
    Dim juego_abridora As String
    Dim juego_tenido As String
    Dim juego_conera As String
    Dim articulo As String
    Dim titulo As String
    Dim hilo As String
    Dim longitud As Double
    Dim conos As String
    Dim cuerda As String
    Dim hinicio As String
    Dim hfinal As String

    Dim id_conerad As String
    Dim codigod As String
    Dim operariod As String
    Dim fechad As Date
    Dim miniciod As Double
    Dim mfinald As Double
    Dim conerad As String
    Public Property gconerad
        Get
            Return conerad
        End Get
        Set(value)
            conerad = value
        End Set
    End Property
    Public Property gmfinald
        Get
            Return mfinald
        End Get
        Set(value)
            mfinald = value
        End Set
    End Property
    Public Property gminiciod
        Get
            Return miniciod
        End Get
        Set(value)
            miniciod = value
        End Set
    End Property
    Public Property gfechad
        Get
            Return fechad
        End Get
        Set(value)
            fechad = value
        End Set
    End Property
    Public Property goperariod
        Get
            Return operariod
        End Get
        Set(value)
            operariod = value
        End Set
    End Property
    Public Property gcodigod
        Get
            Return codigod
        End Get
        Set(value)
            codigod = value
        End Set
    End Property
    Public Property gid_conerad
        Get
            Return id_conerad
        End Get
        Set(value)
            id_conerad = value
        End Set
    End Property
    Public Property ghfinal
        Get
            Return hfinal
        End Get
        Set(value)
            hfinal = value
        End Set
    End Property
    Public Property ghinicio
        Get
            Return hinicio
        End Get
        Set(value)
            hinicio = value
        End Set
    End Property
    Public Property gcuerda
        Get
            Return cuerda
        End Get
        Set(value)
            cuerda = value
        End Set
    End Property
    Public Property gconos
        Get
            Return conos
        End Get
        Set(value)
            conos = value
        End Set
    End Property
    Public Property glongitud
        Get
            Return longitud
        End Get
        Set(value)
            longitud = value
        End Set
    End Property
    Public Property ghilo
        Get
            Return hilo
        End Get
        Set(value)
            hilo = value
        End Set
    End Property
    Public Property gtitulo
        Get
            Return titulo
        End Get
        Set(value)
            titulo = value
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
    Public Property gjuego_conera
        Get
            Return juego_conera
        End Get
        Set(value)
            juego_conera = value
        End Set
    End Property
    Public Property gjuego_tenido
        Get
            Return juego_tenido
        End Get
        Set(value)
            juego_tenido = value
        End Set
    End Property
    Public Property gjuego_abridora
        Get
            Return juego_abridora
        End Get
        Set(value)
            juego_abridora = value
        End Set
    End Property
    Public Property gjuego_urdido
        Get
            Return juego_urdido
        End Get
        Set(value)
            juego_urdido = value
        End Set
    End Property
    Public Property gid_conera
        Get
            Return id_conera
        End Get
        Set(value)
            id_conera = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal id_conera As String, ByVal juego_urdido As String, ByVal juego_abridora As String, ByVal juego_conera As String, ByVal articulo As String, ByVal titulo As String, ByVal hilo As String,
       ByVal juego_tenido As String, ByVal longitud As Double, ByVal conos As String, ByVal cuerda As String, ByVal hinicio As String, ByVal hfinal As String, ByVal id_conerad As String, ByVal codigod As String,
            ByVal operariod As String, ByVal fechad As Date, ByVal miniciod As Double, ByVal mfinald As Double, ByVal conerad As String)
        gid_conera = id_conera
        gjuego_urdido = juego_urdido
        gjuego_abridora = juego_abridora
        gjuego_tenido = juego_tenido
        gjuego_conera = juego_conera
        garticulo = articulo
        gtitulo = titulo
        ghilo = hilo
        glongitud = longitud
        gconos = conos
        gcuerda = cuerda
        ghinicio = hinicio
        ghfinal = hfinal
        gid_conerad = id_conerad
        gcodigod = codigod
        goperariod = operariod
        gfechad = fechad
        gminiciod = miniciod
        gmfinald = mfinald
        gconerad = conerad
    End Sub
End Class
