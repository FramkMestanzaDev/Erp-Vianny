Public Class vresto_op
    Dim ncom_3 As String
    Dim linea_3 As String
    Dim corel As String
    Dim color As String
    Dim ncom As String
    Dim correl As String
    Dim cantidad As Double
    Dim codcom As String
    Dim codtol As String
    Dim ncom4 As String
    Dim correlq As String
    Dim talla As String
    Dim codtal As String
    Dim colorrt As String
    Dim ccia As String
    Dim ratio As String
    Public Property gratio
        Get
            Return ratio
        End Get
        Set(value)
            ratio = value
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
    Public Property gcolorrt
        Get
            Return colorrt
        End Get
        Set(value)
            colorrt = value
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
    Public Property gcorrelq
        Get
            Return correlq
        End Get
        Set(value)
            correlq = value
        End Set
    End Property
    Public Property gtalla
        Get
            Return talla
        End Get
        Set(value)
            talla = value
        End Set
    End Property
    Public Property gcodtal
        Get
            Return codtal
        End Get
        Set(value)
            codtal = value
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
    Public Property gcorrel
        Get
            Return correl
        End Get
        Set(value)
            correl = value
        End Set
    End Property
    Public Property gcantidad
        Get
            Return cantidad
        End Get
        Set(value)
            cantidad = value
        End Set
    End Property
    Public Property gcodcom
        Get
            Return codcom
        End Get
        Set(value)
            codcom = value
        End Set
    End Property
    Public Property gcodtol
        Get
            Return codtol
        End Get
        Set(value)
            codtol = value
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
    Public Property glinea_3
        Get
            Return linea_3
        End Get
        Set(value)
            linea_3 = value
        End Set
    End Property
    Public Property gcorel
        Get
            Return corel
        End Get
        Set(value)
            corel = value
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
    Sub New()

    End Sub
    Sub New(ByVal ncom_3 As String, ByVal linea_3 As String, ByVal corel As String, ByVal color As String, ByVal ccia As String,
            ByVal ncom As String, ByVal correl As String, ByVal cantidad As String, ByVal codcom As String, ByVal codtol As String,
            ByVal ncom4 As String, ByVal correlq As String, ByVal talla As Double, codtal As String, ByVal colorrt As String, ByVal ratio As String)
        gncom_3 = ncom_3
        glinea_3 = linea_3
        gcorel = corel
        gcolor = color
        gncom = ncom
        gcorrel = correl
        gcantidad = cantidad
        gcodcom = codcom
        gcodtol = codtol
        gncom4 = ncom4
        gcorrelq = correlq
        gtalla = talla
        gcodtal = codtal
        gcolorrt = colorrt
        gccia = ccia
        gratio = ratio
    End Sub

End Class
