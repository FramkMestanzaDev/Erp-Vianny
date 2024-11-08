Public Class vcliente

    Dim codigo As String
    Dim T_persona As String
    Dim r_social As String
    Dim ruc As String
    Dim D_fiscal As String
    Dim email As String
    Dim telefono As String
    Dim email_fac As String
    Dim Vendedor As String
    Dim origen As String
    Dim pais As String
    Dim departamento As String
    Dim provincia As String
    Dim distrito As String
    Dim nombres As String
    Dim apaterno As String
    Dim amaterno As String
    Dim t_doc As String
    Dim contacto As String
    Dim ubigeo As String
    Dim COD_CLI As String
    Dim lcredito As Double
    Dim estado As String
    Dim U_NEGOCIO As String
    Dim almacen As String
    Dim ccia As String
    Dim fcom As Date
    Dim fcom_fin As Date


    Public Property gfcom_fin
        Get
            Return fcom_fin
        End Get
        Set(value)
            fcom_fin = value
        End Set
    End Property
    Public Property gfcom
        Get
            Return fcom
        End Get
        Set(value)
            fcom = value
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
    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
        End Set
    End Property
    Public Property gU_NEGOCIO
        Get
            Return U_NEGOCIO
        End Get
        Set(value)
            U_NEGOCIO = value
        End Set
    End Property
    Public Property glcredito
        Get
            Return lcredito
        End Get
        Set(value)
            lcredito = value
        End Set
    End Property
    Public Property gestado
        Get
            Return estado
        End Get
        Set(value)
            estado = value
        End Set
    End Property
    Public Property gCOD_CLI
        Get
            Return COD_CLI
        End Get
        Set(value)
            COD_CLI = value
        End Set
    End Property
    Public Property gubigeo
        Get
            Return ubigeo
        End Get
        Set(value)
            ubigeo = value
        End Set
    End Property
    Public Property gnombres
        Get
            Return nombres
        End Get
        Set(value)
            nombres = value
        End Set
    End Property
    Public Property gapaterno
        Get
            Return apaterno
        End Get
        Set(value)
            apaterno = value
        End Set
    End Property
    Public Property gamaterno
        Get
            Return amaterno
        End Get
        Set(value)
            amaterno = value
        End Set
    End Property
    Public Property gt_doc
        Get
            Return t_doc
        End Get
        Set(value)
            t_doc = value
        End Set
    End Property
    Public Property gcontacto
        Get
            Return contacto
        End Get
        Set(value)
            contacto = value
        End Set
    End Property
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Public Property gT_persona
        Get
            Return T_persona
        End Get
        Set(value)
            T_persona = value
        End Set
    End Property
    Public Property gr_social
        Get
            Return r_social
        End Get
        Set(value)
            r_social = value
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
    Public Property gD_fiscal
        Get
            Return D_fiscal
        End Get
        Set(value)
            D_fiscal = value
        End Set
    End Property
    Public Property gemail
        Get
            Return email
        End Get
        Set(value)
            email = value
        End Set
    End Property
    Public Property gtelefono
        Get
            Return telefono
        End Get
        Set(value)
            telefono = value
        End Set
    End Property
    Public Property gemail_fac
        Get
            Return email_fac
        End Get
        Set(value)
            email_fac = value
        End Set
    End Property
    Public Property gVendedor
        Get
            Return Vendedor
        End Get
        Set(value)
            Vendedor = value
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
    Public Property gpais
        Get
            Return pais
        End Get
        Set(value)
            pais = value
        End Set
    End Property
    Public Property gdepartamento
        Get
            Return departamento
        End Get
        Set(value)
            departamento = value
        End Set
    End Property
    Public Property gprovincia
        Get
            Return provincia
        End Get
        Set(value)
            provincia = value
        End Set
    End Property
    Public Property gdistrito
        Get
            Return distrito
        End Get
        Set(value)
            distrito = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal codigo As String,
        ByVal T_persona As String,
        ByVal r_social As String,
        ByVal ruc As String,
        ByVal D_fiscal As String,
        ByVal email As String,
        ByVal telefono As String,
        ByVal email_fac As String,
        ByVal Vendedor As String,
        ByVal origen As String,
        ByVal pais As String,
        ByVal departamento As String,
        ByVal provincia As String,
        ByVal distrito As String,
             ByVal nombres As String,
              ByVal apaterno As String,
               ByVal amaterno As String,
                ByVal t_doc As String,
                 ByVal contacto As String,
            ByVal ubigeo As String, ByVal almacen As String, ByVal ccia As String,
            ByVal COD_CLI As String, ByVal lcredito As Double, ByVal estado As String, ByVal U_NEGOCIO As String, ByVal fcom As Date, ByVal fcom_fin As Date
   )
        gccia = ccia
        galmacen = almacen
        gcodigo = codigo
        gT_persona = T_persona
        gr_social = r_social
        gruc = ruc
        gD_fiscal = D_fiscal
        gemail = email
        gtelefono = telefono
        gemail_fac = email_fac
        gVendedor = Vendedor
        gorigen = origen
        gpais = pais
        gdepartamento = departamento
        gprovincia = provincia
        gdistrito = distrito
        gnombres = nombres
        gapaterno = apaterno
        gamaterno = amaterno
        gt_doc = t_doc
        gcontacto = contacto
        gubigeo = ubigeo
        gCOD_CLI = COD_CLI
        gestado = estado
        glcredito = lcredito
        gU_NEGOCIO = U_NEGOCIO
        gfcom = fcom
        gfcom_fin = fcom_fin
    End Sub
End Class
