Public Class vop
    Dim cia As String
    Dim ncom_3 As String
    Dim fich_3 As String
    Dim fcom_3 As Date
    Dim mone_3 As Integer
    Dim linea_3 As String
    Dim FCome1_3 As Date
    Dim frequerida_3 As Date
    Dim tcam_3 As Double
    Dim program_3 As String
    Dim flag_3 As Double
    Dim descri_3 As String
    Dim alterno_3 As String
    Dim estilo_3 As String
    Dim pfob_3 As Double
    Dim cantp_3 As Double
    Dim cants_3 As Double
    Dim merchan_3 As String
    Dim broker_3 As String
    Dim tela1_3 As String
    Dim telaprinc_3 As String
    Dim fpago_3 As String
    Dim ffinprod_3 As Date
    Dim observ_3 As String
    Dim cod_color As String
    Dim color As String
    Dim tipo As String
    Dim familia As String
    Dim oobservacion2 As String
    Dim estiloemp_3 As String
    Dim ID_REQUERIMIENTO As Integer
    Dim ttela As String
    Public Property gttela
        Get
            Return ttela
        End Get
        Set(value)
            ttela = value
        End Set
    End Property
    Public Property gID_REQUERIMIENTO
        Get
            Return ID_REQUERIMIENTO
        End Get
        Set(value)
            ID_REQUERIMIENTO = value
        End Set
    End Property
    Public Property gestiloemp_3
        Get
            Return estiloemp_3
        End Get
        Set(value)
            estiloemp_3 = value
        End Set
    End Property
    Public Property goobservacion2
        Get
            Return oobservacion2
        End Get
        Set(value)
            oobservacion2 = value
        End Set
    End Property
    Public Property gfamilia
        Get
            Return familia
        End Get
        Set(value)
            familia = value
        End Set
    End Property
    Public Property gcia
        Get
            Return cia
        End Get
        Set(value)
            cia = value
        End Set
    End Property
    Public Property gcod_color
        Get
            Return cod_color
        End Get
        Set(value)
            cod_color = value
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
    Public Property gdescri_3
        Get
            Return descri_3
        End Get
        Set(value)
            descri_3 = value
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
    Public Property gfich_3
        Get
            Return fich_3
        End Get
        Set(value)
            fich_3 = value
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
    Public Property glinea_3
        Get
            Return linea_3
        End Get
        Set(value)
            linea_3 = value
        End Set
    End Property
    Public Property gFCome1_3
        Get
            Return FCome1_3
        End Get
        Set(value)
            FCome1_3 = value
        End Set
    End Property
    Public Property gfrequerida_3
        Get
            Return frequerida_3
        End Get
        Set(value)
            frequerida_3 = value
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
    Public Property gprogram_3
        Get
            Return program_3
        End Get
        Set(value)
            program_3 = value
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
    Public Property galterno_3
        Get
            Return alterno_3
        End Get
        Set(value)
            alterno_3 = value
        End Set
    End Property
    Public Property gestilo_3
        Get
            Return estilo_3
        End Get
        Set(value)
            estilo_3 = value
        End Set
    End Property
    Public Property gpfob_3
        Get
            Return pfob_3
        End Get
        Set(value)
            pfob_3 = value
        End Set
    End Property
    Public Property gcantp_3
        Get
            Return cantp_3
        End Get
        Set(value)
            cantp_3 = value
        End Set
    End Property
    Public Property gcants_3
        Get
            Return cants_3
        End Get
        Set(value)
            cants_3 = value
        End Set
    End Property
    Public Property gmerchan_3
        Get
            Return merchan_3
        End Get
        Set(value)
            merchan_3 = value
        End Set
    End Property

    Public Property gbroker_3
        Get
            Return broker_3
        End Get
        Set(value)
            broker_3 = value
        End Set
    End Property
    Public Property gtela1_3
        Get
            Return tela1_3
        End Get
        Set(value)
            tela1_3 = value
        End Set
    End Property
    Public Property gtelaprinc_3
        Get
            Return telaprinc_3
        End Get
        Set(value)
            telaprinc_3 = value
        End Set
    End Property
    Public Property gfpago_3
        Get
            Return fpago_3
        End Get
        Set(value)
            fpago_3 = value
        End Set
    End Property
    Public Property gffinprod_3
        Get
            Return ffinprod_3
        End Get
        Set(value)
            ffinprod_3 = value
        End Set
    End Property
    Public Property gobserv_3
        Get
            Return observ_3
        End Get
        Set(value)
            observ_3 = value
        End Set
    End Property
    Public Property gtipo
        Get
            Return tipo
        End Get
        Set(value)
            tipo = value
        End Set
    End Property
    Sub New()

    End Sub

    Sub New(ByVal ncom_3 As String,
    ByVal fich_3 As String,
        ByVal fcom_3 As Date,
        ByVal mone_3 As Integer,
        ByVal linea_3 As String,
        ByVal FCome1_3 As Date,
        ByVal frequerida_3 As Date,
        ByVal tcam_3 As Double,
        ByVal program_3 As String,
        ByVal flag_3 As Double,
        ByVal descri_3 As String,
        ByVal alterno_3 As String,
        ByVal estilo_3 As String,
        ByVal pfob_3 As Double,
        ByVal cantp_3 As Double,
        ByVal cants_3 As Double,
        ByVal merchan_3 As String,
        ByVal broker_3 As String,
        ByVal tela1_3 As String,
        ByVal telaprinc_3 As String,
        ByVal fpago_3 As String, ByVal estiloemp_3 As String,
        ByVal ffinprod_3 As Date, ByVal oobservacion2 As String, ByVal ttela As String,
        ByVal observ_3 As String, ByVal cod_color As String, ByVal color As String, ByVal cia As String, ByVal tipo As String, ByVal familia As String)
        gncom_3 = ncom_3
        gfich_3 = fich_3
        gfcom_3 = fcom_3
        gmone_3 = mone_3
        glinea_3 = linea_3
        gFCome1_3 = FCome1_3
        gfrequerida_3 = frequerida_3
        gtcam_3 = tcam_3
        gprogram_3 = program_3
        gflag_3 = flag_3
        gdescri_3 = descri_3
        galterno_3 = alterno_3
        gestilo_3 = estilo_3
        gpfob_3 = pfob_3
        gcantp_3 = cantp_3
        gcants_3 = cants_3
        gmerchan_3 = merchan_3
        gbroker_3 = broker_3
        gtela1_3 = tela1_3
        gtelaprinc_3 = telaprinc_3
        gfpago_3 = fpago_3
        gffinprod_3 = ffinprod_3
        gobserv_3 = observ_3
        gcod_color = cod_color
        gcolor = color
        gcia = cia
        gtipo = tipo
        gfamilia = familia
        goobservacion2 = oobservacion2
        gestiloemp_3 = estiloemp_3
        gttela = ttela
    End Sub
End Class
