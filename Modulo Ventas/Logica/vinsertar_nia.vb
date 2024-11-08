Public Class vinsertar_nia
    Dim ccia As String
    Dim ncom As String
    Dim fecha As DateTime
    Dim glosa As String
    Dim guia As String
    Dim doc As String
    Dim almacen As String
    Dim origen As String
    Dim int As String
    Dim empresa As String
    Dim tidoc As String
    Dim usuario As String
    Dim ano As String
    Dim fase As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gfase
        Get
            Return fase
        End Get
        Set(value)
            fase = value
        End Set
    End Property
    Public Property gano
        Get
            Return ano
        End Get
        Set(value)
            ano = value
        End Set
    End Property
    Public Property gtidoc
        Get
            Return tidoc
        End Get
        Set(value)
            tidoc = value
        End Set
    End Property
    Public Property gusuario
        Get
            Return usuario
        End Get
        Set(value)
            usuario = value
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
    Public Property gint
        Get
            Return int
        End Get
        Set(value)
            int = value
        End Set
    End Property
    Public Property gorigen
        Get
            Return origen
        End Get
        Set(value)
            origen = value
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
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Public Property gglosa
        Get
            Return glosa
        End Get
        Set(value)
            glosa = value
        End Set
    End Property
    Public Property gguia
        Get
            Return guia
        End Get
        Set(value)
            guia = value
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
    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByValempresa As String, ByVal ncom As Integer, ByVal fecha As DateTime, ByVal fase As String,
            ByVal glosa As String, ByVal doc As String, ByVal guia As String, ByVal almacen As String, ByVal ccia As String,
            ByVal origen As String, ByVal int As String, ByVal tidoc As String, ByVal usuario As String, ByVal ano As String)
        gncom = ncom
        gfecha = fecha
        gglosa = glosa
        gdoc = doc
        gguia = guia
        galmacen = almacen
        gorigen = origen
        gint = int
        gempresa = empresa
        gtidoc = tidoc
        gusuario = usuario
        gano = ano
        gfase = fase
        gccia = ccia
    End Sub
End Class
