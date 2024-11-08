Public Class vmatg040f
    Dim ccia_4 As String
    Dim ncom4 As String
    Dim fecha As Date
    Dim observacion As String
    Dim usuario As String
    Dim partida_4 As String
    Dim color_4 As String
    Dim prvtin_4 As String
    Dim prvtej_4 As String
    Dim acabado_4 As String
    Dim presen_4 As String
    Dim psxrolst_4 As Double
    Dim nrevrol_4 As Double
    Dim rpmrol_4 As Double

    Public Property gprvtin_4
        Get
            Return prvtin_4
        End Get
        Set(value)
            prvtin_4 = value
        End Set
    End Property
    Public Property grpmrol_4
        Get
            Return rpmrol_4
        End Get
        Set(value)
            rpmrol_4 = value
        End Set
    End Property
    Public Property gnrevrol_4
        Get
            Return nrevrol_4
        End Get
        Set(value)
            nrevrol_4 = value
        End Set
    End Property
    Public Property gpsxrolst_4
        Get
            Return psxrolst_4
        End Get
        Set(value)
            psxrolst_4 = value
        End Set
    End Property
    Public Property gpresen_4
        Get
            Return presen_4
        End Get
        Set(value)
            presen_4 = value
        End Set
    End Property
    Public Property gacabado_4
        Get
            Return acabado_4
        End Get
        Set(value)
            acabado_4 = value
        End Set
    End Property
    Public Property gprvtej_4
        Get
            Return prvtej_4
        End Get
        Set(value)
            prvtej_4 = value
        End Set
    End Property


    Public Property gcolor_4
        Get
            Return color_4
        End Get
        Set(value)
            color_4 = value
        End Set
    End Property
    Public Property gpartida_4
        Get
            Return partida_4
        End Get
        Set(value)
            partida_4 = value
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
    Public Property gobservacion
        Get
            Return observacion
        End Get
        Set(value)
            observacion = value
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
    Public Property gncom4
        Get
            Return ncom4
        End Get
        Set(value)
            ncom4 = value
        End Set
    End Property
    Public Property gccia_4
        Get
            Return ccia_4
        End Get
        Set(value)
            ccia_4 = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ccia_4 As String,
    ByVal ncom4 As String,
        ByVal fecha As Date,
        ByVal observacion As String,
        ByVal usuario As String,
        ByVal partida_4 As String,
        ByVal color_4 As String,
        ByVal prvtin_4 As String,
        ByVal prvtej_4 As String,
        ByVal acabado_4 As String,
        ByVal presen_4 As String,
        ByVal psxrolst_4 As Double,
        ByVal nrevrol_4 As Double,
        ByVal rpmrol_4 As Double)
        gccia_4 = ccia_4
        gncom4 = ncom4
        gfecha = fecha
        gobservacion = observacion
        gusuario = usuario
        gpartida_4 = partida_4
        gcolor_4 = color_4
        gprvtin_4 = prvtin_4
        gprvtej_4 = prvtej_4
        gacabado_4 = acabado_4
        gpresen_4 = presen_4
        gpsxrolst_4 = psxrolst_4
        gnrevrol_4 = nrevrol_4
        grpmrol_4 = rpmrol_4
    End Sub
End Class
