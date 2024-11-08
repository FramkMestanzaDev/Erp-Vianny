Public Class fwilder
    Dim op As String
    Dim cotizacion As String
    Dim nia As String
    Dim dir_nia As String
    Dim factura As String
    Dim dir_fac As String
    Dim guia As String
    Dim dir_guia As String
    Dim ocompras As String
    Dim dir_ocompra As String
    Dim lacrado As String
    Dim lacrado_contab As String

    Public Property glacrado_contab
        Get
            Return lacrado_contab
        End Get
        Set(value)
            lacrado_contab = value
        End Set
    End Property
    Public Property gop
        Get
            Return op
        End Get
        Set(value)
            op = value
        End Set
    End Property
    Public Property gcotizacion
        Get
            Return cotizacion
        End Get
        Set(value)
            cotizacion = value
        End Set
    End Property
    Public Property gnia
        Get
            Return nia
        End Get
        Set(value)
            nia = value
        End Set
    End Property
    Public Property gdir_nia
        Get
            Return dir_nia
        End Get
        Set(value)
            dir_nia = value
        End Set
    End Property
    Public Property gfactura
        Get
            Return factura
        End Get
        Set(value)
            factura = value
        End Set
    End Property
    Public Property gdir_fac
        Get
            Return dir_fac
        End Get
        Set(value)
            dir_fac = value
        End Set
    End Property
    Public Property gguia
        Get
            Return guia
        End Get
        Set(value)
            guia = value
        End Set
    End Property
    Public Property gdir_guia
        Get
            Return dir_guia
        End Get
        Set(value)
            dir_guia = value
        End Set
    End Property
    Public Property gocompras
        Get
            Return ocompras
        End Get
        Set(value)
            ocompras = value
        End Set
    End Property
    Public Property gdir_ocompra
        Get
            Return dir_ocompra
        End Get
        Set(value)
            dir_ocompra = value
        End Set
    End Property
    Public Property glacrado
        Get
            Return lacrado
        End Get
        Set(value)
            lacrado = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal op As String, ByVal cotizacion As String, ByVal nia As String, ByVal dir_nia As String, ByVal factura As String, ByVal dir_fac As String, ByVal guia As String, ByVal dir_guia As String, ByVal ocompras As String, ByVal dir_ocompra As String, ByVal lacrado As String, ByVal lacrado_contab As String)
        gop = op
        gcotizacion = cotizacion
        gnia = nia
        gdir_nia = dir_nia
        gfactura = factura
        gdir_fac = dir_fac
        gguia = guia
        gdir_guia = dir_guia
        gocompras = ocompras
        gdir_ocompra = dir_ocompra
        glacrado = lacrado
        glacrado_contab = lacrado_contab
    End Sub

End Class
