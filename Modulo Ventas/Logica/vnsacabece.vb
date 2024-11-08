Public Class vnsacabece
    Dim ncom As String
    Dim fecha As Date
    Dim glosa As String
    Dim doc As String
    Dim guia As String
    Dim sfactu As String
    Dim nfactu As String
    Dim ruc As String
    Dim almacen As String
    Dim usuario As String
    Dim orige_sali As String
    Dim origen As String
    Dim ext As String
    Dim tipointexte As String
    Dim ano As String
    Dim tdocento As String
    Dim fase As String
    Dim ccia As String
    Dim adevol As String
    Dim pedorig_3 As String
    Public Property gpedorig_3
        Get
            Return pedorig_3
        End Get
        Set(value)
            pedorig_3 = value
        End Set
    End Property
    Public Property gadevol
        Get
            Return adevol
        End Get
        Set(value)
            adevol = value
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
    Public Property gtipointexte
        Get
            Return tipointexte
        End Get
        Set(value)
            tipointexte = value
        End Set
    End Property
    Public Property gtdocento
        Get
            Return tdocento
        End Get
        Set(value)
            tdocento = value
        End Set
    End Property
    Public Property gext
        Get
            Return ext
        End Get
        Set(value)
            ext = value
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
    Public Property gusuario
        Get
            Return usuario
        End Get
        Set(value)
            usuario = value
        End Set
    End Property
    Public Property gorige_sali
        Get
            Return orige_sali
        End Get
        Set(value)
            orige_sali = value
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
    Public Property gguia
        Get
            Return guia
        End Get
        Set(value)
            guia = value
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
    Public Property gdoc
        Get
            Return doc
        End Get
        Set(value)
            doc = value
        End Set
    End Property
    Public Property gsfactu
        Get
            Return sfactu
        End Get
        Set(value)
            sfactu = value
        End Set
    End Property
    Public Property gnfactu
        Get
            Return nfactu
        End Get
        Set(value)
            nfactu = value
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
    Sub New(ByVal ncom As String, ByVal fecha As Date, ByVal glosa As String, ByVal doc As String, ByVal fase As String, ByVal adevol As String,
            ByVal guia As String, ByVal sfactu As String, ByVal nfactu As String, ByVal ruc As String, ByVal ano As String,
            ByVal almacen As String, ByVal usuario As String, ByVal orige_sali As String, ByVal origen As String, ByVal ccia As String, ByVal pedorig_3 As String,
            ByVal ext As String, ByVal tdocento As String, ByVal tipointexte As String)
        gncom = ncom
        gfecha = fecha
        gglosa = glosa
        gdoc = doc
        gguia = guia
        gsfactu = sfactu
        gnfactu = nfactu
        gruc = ruc
        galmacen =
        gusuario = usuario
        gorige_sali = orige_sali
        gorigen = origen
        gext = ext
        gtdocento = tdocento
        gtipointexte = tipointexte
        gano = ano
        gfase = fase
        gccia = ccia
        gadevol = adevol
        gpedorig_3 = pedorig_3
    End Sub

End Class
