Public Class vqan0300
    Dim ccia_3n As String
    Dim regis_3n As String
    Dim ncom_3n As String
    Dim fecha As DateTime
    Dim origen_3n As String
    Dim Fich_3n As String
    Dim color_3n As String
    Dim fase_3n As String
    Dim estado_3n As String
    Dim Glosa1_3n As String
    Public Property gGlosa1_3n
        Get
            Return Glosa1_3n
        End Get
        Set(value)
            Glosa1_3n = value
        End Set
    End Property
    Public Property gestado_3n
        Get
            Return estado_3n
        End Get
        Set(value)
            estado_3n = value
        End Set
    End Property
    Public Property gfase_3n
        Get
            Return fase_3n
        End Get
        Set(value)
            fase_3n = value
        End Set
    End Property
    Public Property gcolor_3n
        Get
            Return color_3n
        End Get
        Set(value)
            color_3n = value
        End Set
    End Property
    Public Property gFich_3n
        Get
            Return Fich_3n
        End Get
        Set(value)
            Fich_3n = value
        End Set
    End Property
    Public Property gorigen_3n
        Get
            Return origen_3n
        End Get
        Set(value)
            origen_3n = value
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
    Public Property gncom_3n
        Get
            Return ncom_3n
        End Get
        Set(value)
            ncom_3n = value
        End Set
    End Property
    Public Property gregis_3n
        Get
            Return regis_3n
        End Get
        Set(value)
            regis_3n = value
        End Set
    End Property
    Public Property gccia_3n
        Get
            Return ccia_3n
        End Get
        Set(value)
            ccia_3n = value
        End Set
    End Property
    Sub New()

    End Sub

    Sub New(ByVal ccia_3n As String, ByVal Glosa1_3n As String,
    ByVal regis_3n As String,
        ByVal ncom_3n As String,
        ByVal fecha As DateTime,
        ByVal origen_3n As String,
        ByVal Fich_3n As String,
        ByVal color_3n As String,
        ByVal fase_3n As String,
        ByVal estado_3n As String)
        gccia_3n = ccia_3n
        gregis_3n = regis_3n
        gncom_3n = ncom_3n
        gfecha = fecha
        gorigen_3n = origen_3n
        gFich_3n = Fich_3n
        gcolor_3n = color_3n
        gfase_3n = fase_3n
        gestado_3n = estado_3n
        gGlosa1_3n = Glosa1_3n
    End Sub
End Class
