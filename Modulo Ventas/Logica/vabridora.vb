Public Class vabridora
    Dim id As String
    Dim juego_urdido As String
    Dim juego_abridora As String
    Dim juego_tenido As String
    Dim juego_conera As String
    Dim cuerda As String
    Dim articulo As String
    Dim ball As String
    Dim longitud As Double
    Dim titulo As String
    Dim hilo As String
    Dim codigo As String
    Dim trabajador As String
    Dim lote As String
    Dim flag As String
    Dim programa As String
    Dim maquina As String
    Dim juego_abridorad As String
    Dim cuerdad As String
    Dim tachod As String
    Dim fechad As Date
    Dim maquinad As String
    Dim plegadord As String
    Dim hilod As String
    Dim titulod As String
    Dim loted As String
    Dim despachod As String
    Dim flagd As String
    Dim programad As String
    Dim fecha As Date
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
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
    Public Property gjuego_conera
        Get
            Return juego_conera
        End Get
        Set(value)
            juego_conera = value
        End Set
    End Property
    Public Property gdespachod
        Get
            Return despachod
        End Get
        Set(value)
            despachod = value
        End Set
    End Property
    Public Property gmaquina
        Get
            Return maquina
        End Get
        Set(value)
            maquina = value
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
    Public Property gtachod
        Get
            Return tachod
        End Get
        Set(value)
            tachod = value
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
    Public Property gmaquinad
        Get
            Return maquinad
        End Get
        Set(value)
            maquinad = value
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
    Public Property gplegadord
        Get
            Return plegadord
        End Get
        Set(value)
            plegadord = value
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
    Public Property gjuego_abridorad
        Get
            Return juego_abridorad
        End Get
        Set(value)
            juego_abridorad = value
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
    Public Property glote
        Get
            Return lote
        End Get
        Set(value)
            lote = value
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
    Public Property glongitud
        Get
            Return longitud
        End Get
        Set(value)
            longitud = value
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
    Public Property garticulo
        Get
            Return articulo
        End Get
        Set(value)
            articulo = value
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
    Public Property gjuego_abridora
        Get
            Return juego_abridora
        End Get
        Set(value)
            juego_abridora = value
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
    Public Property gtrabajador
        Get
            Return trabajador
        End Get
        Set(value)
            trabajador = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(
       ByVal juego_tenido As String, ByVal juego_abridora As String, ByVal codigo As String, ByVal trabajador As String, ByVal juego_urdido As String, ByVal longitud As Double, ByVal titulo As String, ByVal articulo As String,
        ByVal hilo As String, ByVal ball As String, ByVal lote As String, ByVal flag As String, ByVal programa As String, ByVal juego_abridorad As String, ByVal cuerdad As String,
ByVal plegadord As String, ByVal fechad As Date, ByVal tachod As String, ByVal titulod As String, ByVal hilod As String, ByVal maquinad As String, ByVal loted As String,
        ByVal flagd As String, ByVal programad As String, ByVal id As String, ByVal maquina As String, ByVal despachod As String, ByVal juego_conera As String, ByVal cuerda As String, ByVal fecha As Date
        )
        gid = id
        gjuego_abridora = juego_abridora
        gcodigo = codigo
        gtrabajador = trabajador
        gjuego_urdido = juego_urdido
        glongitud = longitud
        gtitulo = titulo
        ghilo = hilo
        gjuego_tenido = juego_tenido
        glote = lote
        gflag = flag
        gball = ball
        gprograma = programa
        garticulo = articulo
        gfecha = fecha

        gjuego_abridorad = juego_abridorad
        gcuerdad = cuerdad
        gplegadord = plegadord
        gfechad = fechad
        gtachod = tachod
        gtitulod = titulod
        ghilod = hilod
        gmaquinad = maquinad
        gloted = loted
        gflagd = flagd
        gprogramad = programad
        gmaquina = maquina
        gdespachod = despachod
        gjuego_conera = juego_conera
        gcuerda = cuerda
    End Sub
End Class
