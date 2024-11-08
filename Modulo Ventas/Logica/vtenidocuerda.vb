Public Class vtenidocuerda
    Dim id As String
    Dim juego_urdido As String
    Dim juego_abridora As String
    Dim juego_tenido As String
    Dim juego_conera As String
    Dim articuloAs As String
    Dim titulo As String
    Dim hilo As String
    Dim longitud As Double
    Dim lote As String
    Dim codigo As String
    Dim trabajador As String
    Dim flag As String
    Dim programa As String
    Dim observacion As String
    Dim velocidad As String
    Dim ball As String
    Dim fecha As Date

    Dim juego_tenidod As String
    Dim cuerdad As String
    Dim balld As String
    Dim tachod As String
    Dim titulod As String
    Dim hilod As String
    Dim loted As String
    Dim flagd As String
    Dim programad As String
    Dim observaciond As String
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Public Property gobservaciond
        Get
            Return observaciond
        End Get
        Set(value)
            observaciond = value
        End Set
    End Property
    Public Property gvelocidad
        Get
            Return velocidad
        End Get
        Set(value)
            velocidad = value
        End Set
    End Property
    Public Property gball
        Get
            Return ball
        End Get
        Set(value)
            ball = value
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
    Public Property gprogramad
        Get
            Return programad
        End Get
        Set(value)
            programad = value
        End Set
    End Property
    Public Property gflagd
        Get
            Return flagd
        End Get
        Set(value)
            flagd = value
        End Set
    End Property
    Public Property gloted
        Get
            Return loted
        End Get
        Set(value)
            loted = value
        End Set
    End Property
    Public Property ghilod
        Get
            Return hilod
        End Get
        Set(value)
            hilod = value
        End Set
    End Property
    Public Property gtitulod
        Get
            Return titulod
        End Get
        Set(value)
            titulod = value
        End Set
    End Property
    Public Property gtachod
        Get
            Return tachod
        End Get
        Set(value)
            tachod = value
        End Set
    End Property
    Public Property gballd
        Get
            Return balld
        End Get
        Set(value)
            balld = value
        End Set
    End Property
    Public Property gcuerdad
        Get
            Return cuerdad
        End Get
        Set(value)
            cuerdad = value
        End Set
    End Property
    Public Property gjuego_tenidod
        Get
            Return juego_tenidod
        End Get
        Set(value)
            juego_tenidod = value
        End Set
    End Property
    Public Property gobservacion
        Get
            Return observacion
        End Get
        Set(value)
            observacion = value
        End Set
    End Property
    Public Property gprograma
        Get
            Return programa
        End Get
        Set(value)
            programa = value
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
    Public Property glote
        Get
            Return lote
        End Get
        Set(value)
            lote = value
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

    Public Property garticuloAs
        Get
            Return articuloAs
        End Get
        Set(value)
            articuloAs = value
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
    Sub New()

    End Sub
    Sub New(ByVal juego_urdido As String, ByVal programad As String, ByVal juego_tenidod As String, ByVal cuerdad As String, ByVal balld As String, ByVal tachod As String, ByVal titulod As String, ByVal hilod As String, ByVal loted As String, ByVal flagd As String,
            ByVal juego_abridora As String, ByVal juego_tenido As String, ByVal juego_conera As String, ByVal articuloAs As String, ByVal titulo As String, ByVal hilo As String, ByVal longitud As Double, ByVal lote As String, ByVal fecha As Date,
            ByVal codigo As String, ByVal trabajador As String, ByVal flag As String, ByVal programa As String, ByVal observacion As String, ByVal id As String, ByVal ball As String, ByVal velocidad As String, ByVal observaciond As String)
        gjuego_urdido = juego_urdido
        gjuego_abridora = juego_abridora
        gjuego_tenido = juego_tenido
        gjuego_conera = juego_conera
        garticuloAs = articuloAs
        gtitulo = titulo
        ghilo = hilo
        glongitud = longitud
        glote = lote
        gcodigo = codigo
        gtrabajador = trabajador
        gflag = flag
        gprograma = programa
        gobservacion = observacion
        gid = id
        gball = ball
        gvelocidad = velocidad
        gfecha = fecha

        gjuego_tenidod = juego_tenidod
        gcuerdad = cuerdad
        gballd = balld
        gtachod = tachod
        gtitulod = titulod
        ghilod = hilod
        gloted = loted
        gflagd = flagd
        gprogramad = programad
        gobservaciond = observaciond
    End Sub
End Class
