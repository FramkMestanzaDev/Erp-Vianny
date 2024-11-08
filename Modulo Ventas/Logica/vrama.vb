Public Class vrama
    Dim partida As String
	Dim codigo As String
	Dim orden As String
	Dim po As String
	Dim color As String
	Dim articulo As String
	Dim rollos As String
	Dim peso As Double
	Dim lote As String
	Dim ancho As String
	Dim densidad As String
	Dim flag As String
	Dim peso_total As Double
	Dim observacion As String
	Dim fecha As Date
	Dim fecha_rama As Date
    Dim hora As String
    Dim pasado As String
    Dim id As Integer
    Dim hora2 As String
    Dim PROGRAMA As String

    Dim MOTIVO As String
    Dim COMENTARIO As String
    Dim anchor As Double
    Dim densidadr As Integer
    Dim velocidad As Double
    Dim temperatura As Double
    Dim igsa As String
    Dim salimentacion As Double
    Dim dpesado As String
    Dim maquina As String
    Dim ccia As String

    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
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
    Public Property gsalimentacion
        Get
            Return salimentacion
        End Get
        Set(value)
            salimentacion = value
        End Set
    End Property
    Public Property gdpesado
        Get
            Return dpesado
        End Get
        Set(value)
            dpesado = value
        End Set
    End Property
    Public Property gtemperatura
        Get
            Return temperatura
        End Get
        Set(value)
            temperatura = value
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
    Public Property gigsa
        Get
            Return igsa
        End Get
        Set(value)
            igsa = value
        End Set
    End Property
    Public Property gdensidadr
        Get
            Return densidadr
        End Get
        Set(value)
            densidadr = value
        End Set
    End Property
    Public Property ganchor
        Get
            Return anchor
        End Get
        Set(value)
            anchor = value
        End Set
    End Property
    Public Property gCOMENTARIO
        Get
            Return COMENTARIO
        End Get
        Set(value)
            COMENTARIO = value
        End Set
    End Property
    Public Property gMOTIVO
        Get
            Return MOTIVO
        End Get
        Set(value)
            MOTIVO = value
        End Set
    End Property

    Public Property gPROGRAMA
        Get
            Return PROGRAMA
        End Get
        Set(value)
            PROGRAMA = value
        End Set
    End Property

    Public Property ghora2
        Get
            Return hora2
        End Get
        Set(value)
            hora2 = value
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
    Public Property gpasado
        Get
            Return pasado
        End Get
        Set(value)
            pasado = value
        End Set
    End Property
    Public Property ghora
        Get
            Return hora
        End Get
        Set(value)
            hora = value
        End Set
    End Property
    Public Property gfecha_rama
        Get
            Return fecha_rama
        End Get
        Set(value)
            fecha_rama = value
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
    Public Property gobservacion
        Get
            Return observacion
        End Get
        Set(value)
            observacion = value
        End Set
    End Property
    Public Property gpeso_total
        Get
            Return peso_total
        End Get
        Set(value)
            peso_total = value
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
    Public Property glote
        Get
            Return lote
        End Get
        Set(value)
            lote = value
        End Set
    End Property
    Public Property gpeso
        Get
            Return peso
        End Get
        Set(value)
            peso = value
        End Set
    End Property
    Public Property grollos
        Get
            Return rollos
        End Get
        Set(value)
            rollos = value
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
    Public Property gcolor
        Get
            Return color
        End Get
        Set(value)
            color = value
        End Set
    End Property
    Public Property gpo
        Get
            Return po
        End Get
        Set(value)
            po = value
        End Set
    End Property
    Public Property gorden
        Get
            Return orden
        End Get
        Set(value)
            orden = value
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
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal codigo As String, ByVal orden As String, ByVal po As String, ByVal color As String, ByVal articulo As String, ByVal rollos As String, ByVal peso As Double,
            ByVal lote As String, ByVal ancho As String, ByVal densidad As String, ByVal flag As String, ByVal peso_total As Double, ByVal observacion As String, ByVal fecha As Date, ByVal fecha_rama As Date,
            ByVal hora As String, ByVal pasado As String, ByVal id As Integer, ByVal motivo As String, ByVal comentario As String, ByVal anchor As Double, ByVal densidadr As Integer, ByVal maquina As String,
        ByVal PROGRAMA As String, ByVal igsa As String, ByVal temperatura As Double, ByVal velocidad As Double, ByVal dpesado As String, ByVal sa As Double, ByVal ccia As String)
        gsalimentacion = sa
        gdpesado = dpesado
        gpartida = partida
        gcodigo = codigo
        gorden = orden
        gpo = po
        gcolor = color
        garticulo = articulo
        grollos = rollos
        gpeso = peso
        glote = lote
        gancho = ancho
        gdensidad = densidad
        gflag = flag
        gpeso_total = peso_total
        gobservacion = observacion
        gfecha = fecha
        gfecha_rama = fecha_rama
        ghora = hora
        gpasado = pasado
        gid = id
        gPROGRAMA = PROGRAMA
        gigsa = igsa
        gMOTIVO = motivo
        gCOMENTARIO = comentario
        gdensidadr = densidadr
        ganchor = anchor
        gtemperatura = temperatura
        gvelocidad = velocidad
        gmaquina = maquina
        gccia = ccia
    End Sub
End Class
