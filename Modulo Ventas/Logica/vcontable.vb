Public Class vcontable
    Dim ccia_3 As String
    Dim cperiodo_3 As String
    Dim ccom_3 As String
    Dim ncom_3 As String
    Dim tcom_3 As Double
    Dim mone_3 As Double
    Dim fcom_3 As Date
    Dim tcam_3 As Double
    Dim glosa_3 As String
    Dim glosb_3 As String
    Dim totd1_3 As Double
    Dim totd2_3 As Double
    Dim toth1_3 As Double
    Dim toth2_3 As Double
    Dim cuenb_3 As String
    Dim girado_3 As String
    Dim nombcp_3 As String
    Dim fcoma_3 As Date
    Dim nroa_3 As String
    Dim tmov_3 As String
    Dim difcam_3 As Double
    Dim tasien_3 As Double
    Dim flag_3 As Double
    Dim fcome_3 As Date
    Dim usuario_3 As String
    Dim fupdate_3 As String
    Dim ALMACEN As String
    Dim RUBRO As String
    Public Property gRUBRO
        Get
            Return RUBRO
        End Get
        Set(value)
            RUBRO = value
        End Set
    End Property
    Public Property gALMACEN
        Get
            Return ALMACEN
        End Get
        Set(value)
            ALMACEN = value
        End Set
    End Property
    Public Property gfupdate_3
        Get
            Return fupdate_3
        End Get
        Set(value)
            fupdate_3 = value
        End Set
    End Property
    Public Property gusuario_3
        Get
            Return usuario_3
        End Get
        Set(value)
            usuario_3 = value
        End Set
    End Property
    Public Property gfcome_3
        Get
            Return fcome_3
        End Get
        Set(value)
            fcome_3 = value
        End Set
    End Property
    Public Property gflag_3
        Get
            Return flag_3
        End Get
        Set(value)
            flag_3 = value
        End Set
    End Property
    Public Property gtasien_3
        Get
            Return tasien_3
        End Get
        Set(value)
            tasien_3 = value
        End Set
    End Property
    Public Property gdifcam_3
        Get
            Return difcam_3
        End Get
        Set(value)
            difcam_3 = value
        End Set
    End Property
    Public Property gtmov_3
        Get
            Return tmov_3
        End Get
        Set(value)
            tmov_3 = value
        End Set
    End Property
    Public Property gnroa_3
        Get
            Return nroa_3
        End Get
        Set(value)
            nroa_3 = value
        End Set
    End Property
    Public Property gfcoma_3
        Get
            Return fcoma_3
        End Get
        Set(value)
            fcoma_3 = value
        End Set
    End Property
    Public Property gnombcp_3
        Get
            Return nombcp_3
        End Get
        Set(value)
            nombcp_3 = value
        End Set
    End Property
    Public Property ggirado_3
        Get
            Return girado_3
        End Get
        Set(value)
            girado_3 = value
        End Set
    End Property
    Public Property gcuenb_3
        Get
            Return cuenb_3
        End Get
        Set(value)
            cuenb_3 = value
        End Set
    End Property
    Public Property gtoth2_3
        Get
            Return toth2_3
        End Get
        Set(value)
            toth2_3 = value
        End Set
    End Property
    Public Property gtoth1_3
        Get
            Return toth1_3
        End Get
        Set(value)
            toth1_3 = value
        End Set
    End Property
    Public Property gtotd2_3
        Get
            Return totd2_3
        End Get
        Set(value)
            totd2_3 = value
        End Set
    End Property
    Public Property gtotd1_3
        Get
            Return totd1_3
        End Get
        Set(value)
            totd1_3 = value
        End Set
    End Property
    Public Property gglosb_3
        Get
            Return glosb_3
        End Get
        Set(value)
            glosb_3 = value
        End Set
    End Property
    Public Property gglosa_3
        Get
            Return glosa_3
        End Get
        Set(value)
            glosa_3 = value
        End Set
    End Property
    Public Property gtcam_3
        Get
            Return tcam_3
        End Get
        Set(value)
            tcam_3 = value
        End Set
    End Property
    Public Property gfcom_3
        Get
            Return fcom_3
        End Get
        Set(value)
            fcom_3 = value
        End Set
    End Property
    Public Property gmone_3
        Get
            Return mone_3
        End Get
        Set(value)
            mone_3 = value
        End Set
    End Property
    Public Property gtcom_3
        Get
            Return tcom_3
        End Get
        Set(value)
            tcom_3 = value
        End Set
    End Property
    Public Property gncom_3
        Get
            Return ncom_3
        End Get
        Set(value)
            ncom_3 = value
        End Set
    End Property
    Public Property gccom_3
        Get
            Return ccom_3
        End Get
        Set(value)
            ccom_3 = value
        End Set
    End Property
    Public Property gcperiodo_3
        Get
            Return cperiodo_3
        End Get
        Set(value)
            cperiodo_3 = value
        End Set
    End Property
    Public Property gccia_3
        Get
            Return ccia_3
        End Get
        Set(value)
            ccia_3 = value
        End Set
    End Property
    Sub New()

    End Sub

    Sub New(
            ByVal ccia_3 As String,
    ByVal cperiodo_3 As String,
        ByVal ccom_3 As String,
        ByVal ncom_3 As String,
        ByVal tcom_3 As Double,
        ByVal mone_3 As Double,
        ByVal fcom_3 As Date,
        ByVal tcam_3 As Double,
        ByVal glosa_3 As String,
        ByVal glosb_3 As String,
        ByVal totd1_3 As Double,
        ByVal totd2_3 As Double,
        ByVal toth1_3 As Double,
        ByVal toth2_3 As Double,
        ByVal cuenb_3 As String,
        ByVal girado_3 As String,
        ByVal nombcp_3 As String,
        ByVal fcoma_3 As Date,
        ByVal nroa_3 As String,
        ByVal tmov_3 As String,
        ByVal difcam_3 As Double,
        ByVal tasien_3 As Double,
        ByVal flag_3 As Double,
        ByVal fcome_3 As Date,
        ByVal usuario_3 As String,
        ByVal fupdate_3 As String,
            ByVal ALMACEN As String,
            ByVal RUBRO As String
   )
        gccia_3 = ccia_3
        gcperiodo_3 = cperiodo_3
        gccom_3 = ccom_3
        gncom_3 = ncom_3
        gtcom_3 = tcom_3
        gmone_3 = mone_3
        gfcom_3 = fcom_3
        gtcam_3 = tcam_3
        gglosa_3 = glosa_3
        gglosb_3 = glosb_3
        gtotd1_3 = totd1_3
        gtotd2_3 = totd2_3
        gtoth1_3 = toth1_3
        gtoth2_3 = toth2_3
        gcuenb_3 = cuenb_3
        ggirado_3 = girado_3
        gnombcp_3 = nombcp_3
        gfcoma_3 = fcoma_3
        gnroa_3 = nroa_3
        gtmov_3 = tmov_3
        gdifcam_3 = difcam_3
        gtasien_3 = tasien_3
        gflag_3 = flag_3
        gfcome_3 = fcome_3
        gusuario_3 = usuario_3
        gfupdate_3 = fupdate_3
        gALMACEN = ALMACEN
        gRUBRO = RUBRO
    End Sub
End Class
