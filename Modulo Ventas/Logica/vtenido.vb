Public Class vtenido
    Dim ntenido As String
    Dim hora As String
    Dim ph As Double
    Dim mv As Double
    Dim indigo As Double
    Dim hidro As Double
    Dim be As Double
    Dim veloc As Double
    Dim tinas As Double
    Dim ds_colorant As Double
    Dim ds_hidrosuf As Double
    Dim ds_soda As Double
    Dim ds_pquimic As Double
    Dim ds_neutral As Double
    Dim ds_fijad As Double
    Dim ph_neutral As Double
    Dim ph_fijad As Double
    Dim ph_suaviz As Double
    Dim sol As Double
    Dim observaciones As String

    Dim fecha As Date
    Dim metraje As Double
    Dim articulo As String
    Dim lote As String
    Dim nhoja As String
    Dim estado As String
    Dim estado1 As String
    Public Property gestado
        Get
            Return estado
        End Get
        Set(value)
            estado = value
        End Set
    End Property
    Public Property gestado1
        Get
            Return estado1
        End Get
        Set(value)
            estado1 = value
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
    Public Property gnhoja
        Get
            Return nhoja
        End Get
        Set(value)
            nhoja = value
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
    Public Property garticulo
        Get
            Return articulo
        End Get
        Set(value)
            articulo = value
        End Set
    End Property
    Public Property gmetraje
        Get
            Return metraje
        End Get
        Set(value)
            metraje = value
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
    Public Property gobservaciones
        Get
            Return observaciones
        End Get
        Set(value)
            observaciones = value
        End Set
    End Property
    Public Property gsol
        Get
            Return sol
        End Get
        Set(value)
            sol = value
        End Set
    End Property
    Public Property gph_suaviz
        Get
            Return ph_suaviz
        End Get
        Set(value)
            ph_suaviz = value
        End Set
    End Property
    Public Property gph_fijad
        Get
            Return ph_fijad
        End Get
        Set(value)
            ph_fijad = value
        End Set
    End Property
    Public Property gph_neutral
        Get
            Return ph_neutral
        End Get
        Set(value)
            ph_neutral = value
        End Set
    End Property
    Public Property gds_fijad
        Get
            Return ds_fijad
        End Get
        Set(value)
            ds_fijad = value
        End Set
    End Property
    Public Property gds_neutral
        Get
            Return ds_neutral
        End Get
        Set(value)
            ds_neutral = value
        End Set
    End Property
    Public Property gds_pquimic
        Get
            Return ds_pquimic
        End Get
        Set(value)
            ds_pquimic = value
        End Set
    End Property
    Public Property gds_soda
        Get
            Return ds_soda
        End Get
        Set(value)
            ds_soda = value
        End Set
    End Property
    Public Property gds_hidrosuf
        Get
            Return ds_hidrosuf
        End Get
        Set(value)
            ds_hidrosuf = value
        End Set
    End Property
    Public Property gds_colorant
        Get
            Return ds_colorant
        End Get
        Set(value)
            ds_colorant = value
        End Set
    End Property
    Public Property gtinas
        Get
            Return tinas
        End Get
        Set(value)
            tinas = value
        End Set
    End Property
    Public Property gveloc
        Get
            Return veloc
        End Get
        Set(value)
            veloc = value
        End Set
    End Property
    Public Property gbe
        Get
            Return be
        End Get
        Set(value)
            be = value
        End Set
    End Property
    Public Property ghidro
        Get
            Return hidro
        End Get
        Set(value)
            hidro = value
        End Set
    End Property
    Public Property gindigo
        Get
            Return indigo
        End Get
        Set(value)
            indigo = value
        End Set
    End Property
    Public Property gmv
        Get
            Return mv
        End Get
        Set(value)
            mv = value
        End Set
    End Property
    Public Property gph
        Get
            Return ph
        End Get
        Set(value)
            ph = value
        End Set
    End Property
    Public Property gntenido
        Get
            Return ntenido
        End Get
        Set(value)
            ntenido = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal ntenido As String,
            ByVal hora As String,
    ByVal ph As Double,
        ByVal mv As Double,
        ByVal indigo As Double,
        ByVal hidro As Double,
        ByVal be As Double,
        ByVal veloc As Double,
        ByVal tinas As Double,
        ByVal ds_colorant As Double,
        ByVal ds_hidrosuf As Double,
        ByVal ds_soda As Double,
        ByVal ds_pquimic As Double,
        ByVal ds_neutral As Double,
        ByVal ds_fijad As Double,
        ByVal ph_neutral As Double,
        ByVal ph_fijad As Double,
        ByVal ph_suaviz As Double,
        ByVal sol As Double,
        ByVal observaciones As String,
             ByVal fecha As Date,
    ByVal metraje As String,
        ByVal articulo As String,
        ByVal lote As String,
        ByVal nhoja As String, ByVal estado As String, ByVal estado1 As String)
        gntenido = ntenido
        ghora = hora
        gph = ph
        gmv = mv
        gindigo = indigo
        ghidro = hidro
        gbe = be
        gveloc = veloc
        gtinas = tinas
        gds_colorant = ds_colorant
        gds_hidrosuf = ds_hidrosuf
        gds_soda = ds_soda
        gds_pquimic = ds_pquimic
        gds_neutral = ds_neutral
        gds_fijad = ds_fijad
        gph_neutral = ph_neutral
        gph_fijad = ph_fijad
        gph_suaviz = ph_suaviz
        gsol = sol
        gobservaciones = observaciones
        gfecha = fecha
        gmetraje = metraje
        garticulo = articulo
        glote = lote
        gnhoja = nhoja
        gestado1 = estado1
        gestado = estado
    End Sub
End Class
