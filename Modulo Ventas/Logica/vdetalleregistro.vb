Public Class vdetalleregistro

    Dim cperiodo_3a As String
    Dim ncom_3a As String
    Dim item_3a As String
    Dim linea_3a As String
    Dim producto As String
    Dim cant_3a As Double
    Dim unid_3a As String
    Dim preun1_3a As Double
    Dim pvta1_3a As Double
    Dim vvta1_3a As Double
    Dim igv1_3a As Double
    Dim tot1_3a As Double
    Dim preun2_3a As Double
    Dim pvta2_3a As Double
    Dim vvta2_3a As Double
    Dim igv2_3a As Double
    Dim tot2_3a As Double
    Dim obser1_3a As String
    Dim porven_3a As Double
    Dim ordenp_3a As String
    Dim grm_3a As String
    Dim flag_3a As String
    Dim partida As String
    Dim empresa As String
    Public Property gempresa
        Get
            Return empresa
        End Get
        Set(value)
            empresa = value
        End Set
    End Property
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
        End Set
    End Property
    Public Property gcperiodo_3a
        Get
            Return cperiodo_3a
        End Get
        Set(value)
            cperiodo_3a = value
        End Set
    End Property
    Public Property gncom_3a
        Get
            Return ncom_3a
        End Get
        Set(value)
            ncom_3a = value
        End Set
    End Property
    Public Property gitem_3a
        Get
            Return item_3a
        End Get
        Set(value)
            item_3a = value
        End Set
    End Property
    Public Property glinea_3a
        Get
            Return linea_3a
        End Get
        Set(value)
            linea_3a = value
        End Set
    End Property
    Public Property gproducto
        Get
            Return producto
        End Get
        Set(value)
            producto = value
        End Set
    End Property
    Public Property gcant_3a
        Get
            Return cant_3a
        End Get
        Set(value)
            cant_3a = value
        End Set
    End Property
    Public Property gunid_3a
        Get
            Return unid_3a
        End Get
        Set(value)
            unid_3a = value
        End Set
    End Property
    Public Property gpreun1_3a
        Get
            Return preun1_3a
        End Get
        Set(value)
            preun1_3a = value
        End Set
    End Property
    Public Property gpvta1_3a
        Get
            Return pvta1_3a
        End Get
        Set(value)
            pvta1_3a = value
        End Set
    End Property
    Public Property gvvta1_3a
        Get
            Return vvta1_3a
        End Get
        Set(value)
            vvta1_3a = value
        End Set
    End Property
    Public Property gigv1_3a
        Get
            Return igv1_3a
        End Get
        Set(value)
            igv1_3a = value
        End Set
    End Property
    Public Property gtot1_3a
        Get
            Return tot1_3a
        End Get
        Set(value)
            tot1_3a = value
        End Set
    End Property
    Public Property gpreun2_3a
        Get
            Return preun2_3a
        End Get
        Set(value)
            preun2_3a = value
        End Set
    End Property
    Public Property gpvta2_3a
        Get
            Return pvta2_3a
        End Get
        Set(value)
            pvta2_3a = value
        End Set
    End Property
    Public Property gvvta2_3a
        Get
            Return vvta2_3a
        End Get
        Set(value)
            vvta2_3a = value
        End Set
    End Property
    Public Property gigv2_3a
        Get
            Return igv2_3a
        End Get
        Set(value)
            igv2_3a = value
        End Set
    End Property
    Public Property gtot2_3a
        Get
            Return tot2_3a
        End Get
        Set(value)
            tot2_3a = value
        End Set
    End Property
    Public Property gobser1_3a
        Get
            Return obser1_3a
        End Get
        Set(value)
            obser1_3a = value
        End Set
    End Property
    Public Property gporven_3a
        Get
            Return porven_3a
        End Get
        Set(value)
            porven_3a = value
        End Set
    End Property
    Public Property gordenp_3a
        Get
            Return ordenp_3a
        End Get
        Set(value)
            ordenp_3a = value
        End Set
    End Property
    Public Property ggrm_3a
        Get
            Return grm_3a
        End Get
        Set(value)
            grm_3a = value
        End Set
    End Property
    Public Property gflag_3a
        Get
            Return flag_3a
        End Get
        Set(value)
            flag_3a = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal cperiodo_3a As String, ByVal ncom_3a As String, ByVal item_3a As String, ByVal linea_3a As String,
        ByVal producto As String, ByVal cant_3a As Double, ByVal unid_3a As String, ByVal preun1_3a As Double,
        ByVal pvta1_3a As Double, ByVal vvta1_3a As Double, ByVal igv1_3a As Double, ByVal tot1_3a As Double,
        ByVal preun2_3a As Double, ByVal pvta2_3a As Double, ByVal vvta2_3a As Double, ByVal igv2_3a As Double,
        ByVal tot2_3a As Double, ByVal obser1_3a As String, ByVal porven_3a As Double, ByVal ordenp_3a As String, ByVal empresa As String,
        ByVal grm_3a As String, ByVal flag_3a As String, ByVal partida As String
   )
        gcperiodo_3a = cperiodo_3a
        gncom_3a = ncom_3a
        gitem_3a = item_3a
        glinea_3a = linea_3a
        gproducto = producto
        gcant_3a = cant_3a
        gunid_3a = unid_3a
        gpreun1_3a = preun1_3a
        gpvta1_3a = pvta1_3a
        gvvta1_3a = vvta1_3a
        gigv1_3a = igv1_3a
        gtot1_3a = tot1_3a
        gpreun2_3a = preun2_3a
        gpvta2_3a = pvta2_3a
        gvvta2_3a = vvta2_3a
        gigv2_3a = igv2_3a
        gtot2_3a = tot2_3a
        gobser1_3a = obser1_3a
        gporven_3a = porven_3a
        gordenp_3a = ordenp_3a
        ggrm_3a = grm_3a
        gflag_3a = flag_3a
        gpartida = partida
        gempresa = empresa
    End Sub
End Class
