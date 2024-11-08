Public Class vmatp040f
    Dim ccia_4a As String
    Dim ncom_4a As String
    Dim TELA_4A As String
    Dim DIAMET_4A As Double
    Dim GALGA_4A As Double
    Dim CANTREQ_4A As Double
    Dim MAQUINA_4A As String
    Dim ANCHO_4A As Double
    Dim LM1_41A As Double
    Dim LM2_41A As Double
    Dim LM3_41A As Double
    Dim LM4_41A As Double
    Dim LM5_41A As Double
    Public Property gLM5_41A
        Get
            Return LM5_41A
        End Get
        Set(value)
            LM5_41A = value
        End Set
    End Property
    Public Property gLM4_41A
        Get
            Return LM4_41A
        End Get
        Set(value)
            LM4_41A = value
        End Set
    End Property
    Public Property gLM3_41A
        Get
            Return LM3_41A
        End Get
        Set(value)
            LM3_41A = value
        End Set
    End Property
    Public Property gLM2_41A
        Get
            Return LM2_41A
        End Get
        Set(value)
            LM2_41A = value
        End Set
    End Property
    Public Property gLM1_41A
        Get
            Return LM1_41A
        End Get
        Set(value)
            LM1_41A = value
        End Set
    End Property
    Public Property gANCHO_4A
        Get
            Return ANCHO_4A
        End Get
        Set(value)
            ANCHO_4A = value
        End Set
    End Property
    Public Property gMAQUINA_4A
        Get
            Return MAQUINA_4A
        End Get
        Set(value)
            MAQUINA_4A = value
        End Set
    End Property
    Public Property gCANTREQ_4A
        Get
            Return CANTREQ_4A
        End Get
        Set(value)
            CANTREQ_4A = value
        End Set
    End Property
    Public Property gGALGA_4A
        Get
            Return GALGA_4A
        End Get
        Set(value)
            GALGA_4A = value
        End Set
    End Property
    Public Property gDIAMET_4A
        Get
            Return DIAMET_4A
        End Get
        Set(value)
            DIAMET_4A = value
        End Set
    End Property
    Public Property gTELA_4A
        Get
            Return TELA_4A
        End Get
        Set(value)
            TELA_4A = value
        End Set
    End Property
    Public Property gncom_4a
        Get
            Return ncom_4a
        End Get
        Set(value)
            ncom_4a = value
        End Set
    End Property
    Public Property gccia_4a
        Get
            Return ccia_4a
        End Get
        Set(value)
            ccia_4a = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ccia_4a As String,
    ByVal ncom_4a As String,
        ByVal TELA_4A As String,
        ByVal DIAMET_4A As Double,
        ByVal GALGA_4A As Double,
     ByVal CANTREQ_4A As Double,
        ByVal MAQUINA_4A As String,
        ByVal ANCHO_4A As Double,
        ByVal LM1_41A As Double,
        ByVal LM2_41A As Double,
        ByVal LM3_41A As Double,
        ByVal LM4_41A As Double,
        ByVal LM5_41A As Double)
        gccia_4a = ccia_4a
        gncom_4a = ncom_4a
        gTELA_4A = TELA_4A
        gDIAMET_4A = DIAMET_4A
        gGALGA_4A = GALGA_4A
        gCANTREQ_4A = CANTREQ_4A
        gMAQUINA_4A = MAQUINA_4A
        gANCHO_4A = ANCHO_4A
        gLM1_41A = LM1_41A
        gLM2_41A = LM2_41A
        gLM3_41A = LM3_41A
        gLM4_41A = LM4_41A
        gLM5_41A = LM5_41A
    End Sub

End Class
