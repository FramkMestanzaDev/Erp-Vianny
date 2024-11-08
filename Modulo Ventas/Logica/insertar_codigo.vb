Public Class insertar_codigo
    Dim linea_17 As String
    Dim nomb_17 As String
    Dim unid_17 As String
    Dim familia_17 As String
    Dim subfam_17 As String
    Dim tallaje_17 As String
    Dim origen_17 As String
    Dim idcolor_17 As String
    Dim articulo_17 As String
    Dim dmarca_17 As String
    Dim codprod_17 As String
    Dim ncolor_17 As String
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property glinea_17
        Get
            Return linea_17
        End Get
        Set(value)
            linea_17 = value
        End Set
    End Property
    Public Property gnomb_17
        Get
            Return nomb_17
        End Get
        Set(value)
            nomb_17 = value
        End Set
    End Property
    Public Property gunid_17
        Get
            Return unid_17
        End Get
        Set(value)
            unid_17 = value
        End Set
    End Property
    Public Property gfamilia_17
        Get
            Return familia_17
        End Get
        Set(value)
            familia_17 = value
        End Set
    End Property
    Public Property gsubfam_17
        Get
            Return subfam_17
        End Get
        Set(value)
            subfam_17 = value
        End Set
    End Property
    Public Property gtallaje_17
        Get
            Return tallaje_17
        End Get
        Set(value)
            tallaje_17 = value
        End Set
    End Property
    Public Property gorigen_17
        Get
            Return origen_17
        End Get
        Set(value)
            origen_17 = value
        End Set
    End Property
    Public Property gidcolor_17
        Get
            Return idcolor_17
        End Get
        Set(value)
            idcolor_17 = value
        End Set
    End Property
    Public Property garticulo_17
        Get
            Return articulo_17
        End Get
        Set(value)
            articulo_17 = value
        End Set
    End Property
    Public Property gdmarca_17
        Get
            Return dmarca_17
        End Get
        Set(value)
            dmarca_17 = value
        End Set
    End Property
    Public Property gcodprod_17
        Get
            Return codprod_17
        End Get
        Set(value)
            codprod_17 = value
        End Set
    End Property
    Public Property gncolor_17
        Get
            Return ncolor_17
        End Get
        Set(value)
            ncolor_17 = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal linea_17 As String, ByVal nomb_17 As String, ByVal unid_17 As String, ByVal familia_17 As String, ByVal subfam_17 As String, ByVal tallaje_17 As String, ByVal ccia As String,
        ByVal origen_17 As String, ByVal idcolor_17 As String, ByVal articulo_17 As String, ByVal dmarca_17 As String, ByVal codprod_17 As String, ByVal ncolor_17 As String)
        glinea_17 = linea_17
        gnomb_17 = nomb_17
        gunid_17 = unid_17
        gfamilia_17 = familia_17
        gsubfam_17 = subfam_17
        gtallaje_17 = tallaje_17
        gorigen_17 = origen_17
        gidcolor_17 = idcolor_17
        garticulo_17 = articulo_17
        gdmarca_17 = dmarca_17
        gcodprod_17 = codprod_17
        gncolor_17 = ncolor_17
        gccia = ccia
    End Sub
End Class
