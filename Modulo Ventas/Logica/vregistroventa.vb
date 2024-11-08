Public Class vregistroventa
    Dim ruc As String
    Dim cperiodo_3 As String
    Dim ncom_3 As String
    Dim tidoc_3 As String
    Dim sfactu_3 As String
    Dim nfactu_3 As String
    Dim fcom_3 As DateTime
    Dim condp_3 As String
    Dim fich_3 As String
    Dim nomb_3 As String
    Dim mone_3 As String
    Dim tcam_3 As Double
    Dim pvta1_3 As Double
    Dim vvta1_3 As Double
    Dim igv1_3 As Double
    Dim tot1_3 As Double
    Dim pvta2_3 As Double
    Dim vvta2_3 As Double
    Dim igv2_3 As Double
    Dim tot2_3 As Double
    Dim gloa_3 As String
    Dim aigv_3 As Double
    Dim iigv_3 As Double
    Dim anal1_3 As String
    Dim porven_3 As Double
    Dim flag_3 As String
    Dim ocompra_3 As String
    Dim vendedor_3 As String
    Dim porcigv_3 As Double
    Dim tipo_venta As String
    Dim fechacan As Date
    Dim adeelanto As Double
    Dim fecha_cpago As Date
    Dim observacion As String
    Dim empresa As String
    Public Property gempresa
        Get
            Return empresa
        End Get
        Set(value)
            empresa = value
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
    Public Property gfecha_cpago
        Get
            Return fecha_cpago
        End Get
        Set(value)
            fecha_cpago = value
        End Set
    End Property
    Public Property gadeelanto
        Get
            Return adeelanto
        End Get
        Set(value)
            adeelanto = value
        End Set
    End Property
    Public Property gfechacan
        Get
            Return fechacan
        End Get
        Set(value)
            fechacan = value
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
    Public Property gcperiodo_3
        Get
            Return cperiodo_3
        End Get
        Set(value)
            cperiodo_3 = value
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
    Public Property gtidoc_3
        Get
            Return tidoc_3
        End Get
        Set(value)
            tidoc_3 = value
        End Set
    End Property
    Public Property gsfactu_3
        Get
            Return sfactu_3
        End Get
        Set(value)
            sfactu_3 = value
        End Set
    End Property
    Public Property gnfactu_3
        Get
            Return nfactu_3
        End Get
        Set(value)
            nfactu_3 = value
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
    Public Property gcondp_3
        Get
            Return condp_3
        End Get
        Set(value)
            condp_3 = value
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
    Public Property gnomb_3
        Get
            Return nomb_3
        End Get
        Set(value)
            nomb_3 = value
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
    Public Property gtcam_3
        Get
            Return tcam_3
        End Get
        Set(value)
            tcam_3 = value
        End Set
    End Property
    Public Property gpvta1_3
        Get
            Return pvta1_3
        End Get
        Set(value)
            pvta1_3 = value
        End Set
    End Property
    Public Property gvvta1_3
        Get
            Return vvta1_3
        End Get
        Set(value)
            vvta1_3 = value
        End Set
    End Property
    Public Property gigv1_3
        Get
            Return igv1_3
        End Get
        Set(value)
            igv1_3 = value
        End Set
    End Property
    Public Property gtot1_3
        Get
            Return tot1_3
        End Get
        Set(value)
            tot1_3 = value
        End Set
    End Property
    Public Property gpvta2_3
        Get
            Return pvta2_3
        End Get
        Set(value)
            pvta2_3 = value
        End Set
    End Property
    Public Property gvvta2_3
        Get
            Return vvta2_3
        End Get
        Set(value)
            vvta2_3 = value
        End Set
    End Property
    Public Property gtot2_3
        Get
            Return tot2_3
        End Get
        Set(value)
            tot2_3 = value
        End Set
    End Property
    Public Property gigv2_3
        Get
            Return igv2_3
        End Get
        Set(value)
            igv2_3 = value
        End Set
    End Property
    Public Property ggloa_3
        Get
            Return gloa_3
        End Get
        Set(value)
            gloa_3 = value
        End Set
    End Property
    Public Property gaigv_3
        Get
            Return aigv_3
        End Get
        Set(value)
            aigv_3 = value
        End Set
    End Property
    Public Property giigv_3
        Get
            Return iigv_3
        End Get
        Set(value)
            iigv_3 = value
        End Set
    End Property
    Public Property ganal1_3
        Get
            Return anal1_3
        End Get
        Set(value)
            anal1_3 = value
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
    Public Property gocompra_3
        Get
            Return ocompra_3
        End Get
        Set(value)
            ocompra_3 = value
        End Set
    End Property
    Public Property gvendedor_3
        Get
            Return vendedor_3
        End Get
        Set(value)
            vendedor_3 = value
        End Set
    End Property
    Public Property gporven_3
        Get
            Return porven_3
        End Get
        Set(value)
            porven_3 = value
        End Set
    End Property
    Public Property gtipo_venta
        Get
            Return tipo_venta
        End Get
        Set(value)
            tipo_venta = value
        End Set
    End Property
    Public Property gporcigv_3
        Get
            Return porcigv_3
        End Get
        Set(value)
            porcigv_3 = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ruc As String, ByVal cperiodo_3 As String, ByVal ncom_3 As String, ByVal tidoc_3 As String,
        ByVal sfactu_3 As String, ByVal nfactu_3 As String, ByVal fcom_3 As DateTime, ByVal condp_3 As String,
        ByVal fich_3 As String, ByVal nomb_3 As String, ByVal mone_3 As String, ByVal tcam_3 As Double, ByVal pvta1_3 As Double,
        ByVal vvta1_3 As Double, ByVal igv1_3 As Double, ByVal tot1_3 As Double, ByVal pvta2_3 As Double,
        ByVal vvta2_3 As Double, ByVal igv2_3 As Double, ByVal tot2_3 As Double, ByVal gloa_3 As String,
        ByVal aigv_3 As Double, ByVal iigv_3 As Double, ByVal anal1_3 As String, ByVal porven_3 As Double,
        ByVal flag_3 As String, ByVal ocompra_3 As String, ByVal vendedor_3 As String, ByVal porcigv_3 As Double, ByVal empresa As String,
        ByVal tipo_venta As String, ByVal fechacan As Date, ByVal adeelanto As Double, ByVal fecha_cpago As Date, ByVal observacion As String)
        gruc = ruc
        gcperiodo_3 = cperiodo_3
        gncom_3 = ncom_3
        gtidoc_3 = tidoc_3
        gsfactu_3 = sfactu_3
        gnfactu_3 = nfactu_3
        gfcom_3 = fcom_3
        gcondp_3 = condp_3
        gfich_3 = fich_3
        gnomb_3 = nomb_3
        gmone_3 = mone_3
        gtcam_3 = tcam_3
        gpvta1_3 = pvta1_3
        gvvta1_3 = vvta1_3
        gigv1_3 = igv1_3
        gtot1_3 = tot1_3
        gpvta2_3 = pvta2_3
        gvvta2_3 = vvta2_3
        gigv2_3 = igv2_3
        gtot2_3 = tot2_3
        ggloa_3 = gloa_3
        gaigv_3 = aigv_3
        giigv_3 = iigv_3
        ganal1_3 = anal1_3
        gporven_3 = porven_3
        gflag_3 = flag_3
        gocompra_3 = ocompra_3
        gvendedor_3 = vendedor_3
        gporcigv_3 = porcigv_3
        gtipo_venta = tipo_venta
        gadeelanto = adeelanto
        gfechacan = fechacan
        gfecha_cpago = fecha_cpago
        gobservacion = observacion
        gempresa = empresa
    End Sub
End Class
