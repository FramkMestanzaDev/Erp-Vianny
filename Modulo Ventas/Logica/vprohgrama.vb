Public Class vprohgrama
    Dim programa As String
    Dim cargadas As String
    Dim fecha As Date
    Dim kg As Double
    Dim flag As String
    Dim abridorad As String
    Dim urdidod As String
    Dim conerad As String
    Dim tenidod As String
    Dim programad As String
    Dim flagd As String
    Dim articulo As String
    Dim titulo As String
    Dim hilo As String

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
    Public Property gflagd
        Get
            Return flagd
        End Get
        Set(value)
            flagd = value
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
    Public Property gprogramad
        Get
            Return programad
        End Get
        Set(value)
            programad = value
        End Set
    End Property
    Public Property gtenidod
        Get
            Return tenidod
        End Get
        Set(value)
            tenidod = value
        End Set
    End Property
    Public Property gconerad
        Get
            Return conerad
        End Get
        Set(value)
            conerad = value
        End Set
    End Property
    Public Property gurdidod
        Get
            Return urdidod
        End Get
        Set(value)
            urdidod = value
        End Set
    End Property
    Public Property gabridorad
        Get
            Return abridorad
        End Get
        Set(value)
            abridorad = value
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
    Public Property gcargadas
        Get
            Return cargadas
        End Get
        Set(value)
            cargadas = value
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
    Public Property gkg
        Get
            Return kg
        End Get
        Set(value)
            kg = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal programa As String, ByVal fecha As Date, ByVal kg As Double, ByVal abridorad As String, ByVal urdidod As String, ByVal conerad As String, ByVal tenidod As String, ByVal programad As String,
        ByVal cargadas As String, ByRef flag As String, ByVal flagd As String, ByVal articulo As String, ByVal titulo As String, ByVal hilo As String)
        gprograma = programa
        gcargadas = cargadas
        gfecha = fecha
        gkg = kg
        gflag = flag
        gabridorad = abridorad
        gurdidod = urdidod
        gconerad = conerad
        gtenidod = tenidod
        gprogramad = programad
        gflagd = flagd
        garticulo = articulo
        gtitulo = titulo
        ghilo = hilo
    End Sub
End Class
