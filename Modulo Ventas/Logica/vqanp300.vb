Public Class vqanp300
    Dim ccia_3p As String
    Dim regis_3p As String
    Dim Op_3p As String
    Dim linea_3p As String
    Dim galga_3p As Double
    Dim Ancho_3p As Double
    Dim kgreq_3p As Double
    Dim observ_3p As String
    Public Property gobserv_3p
        Get
            Return observ_3p
        End Get
        Set(value)
            observ_3p = value
        End Set
    End Property
    Public Property gkgreq_3p
        Get
            Return kgreq_3p
        End Get
        Set(value)
            kgreq_3p = value
        End Set
    End Property
    Public Property gAncho_3p
        Get
            Return Ancho_3p
        End Get
        Set(value)
            Ancho_3p = value
        End Set
    End Property
    Public Property ggalga_3p
        Get
            Return galga_3p
        End Get
        Set(value)
            galga_3p = value
        End Set
    End Property
    Public Property glinea_3p
        Get
            Return linea_3p
        End Get
        Set(value)
            linea_3p = value
        End Set
    End Property
    Public Property gOp_3p
        Get
            Return Op_3p
        End Get
        Set(value)
            Op_3p = value
        End Set
    End Property
    Public Property gregis_3p
        Get
            Return regis_3p
        End Get
        Set(value)
            regis_3p = value
        End Set
    End Property
    Public Property gccia_3p
        Get
            Return ccia_3p
        End Get
        Set(value)
            ccia_3p = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ccia_3p As String,
    ByVal regis_3p As String,
        ByVal Op_3p As String,
        ByVal linea_3p As String,
        ByVal galga_3p As Double,
        ByVal Ancho_3p As Double,
        ByVal kgreq_3p As Double,
        ByVal observ_3p As String)
        gccia_3p = ccia_3p
        gregis_3p = regis_3p
        gOp_3p = Op_3p
        glinea_3p = linea_3p
        ggalga_3p = galga_3p
        gAncho_3p = Ancho_3p
        gkgreq_3p = kgreq_3p
        gobserv_3p = observ_3p
    End Sub
End Class
