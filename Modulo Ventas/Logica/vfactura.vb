Public Class vfactura
    Dim sfactu As String
    Dim nfactu As String
    Dim periodo As String
    Dim MONTO As Double
    Dim vendedor As String
    Dim mes As String
    Dim ndoc As String
    Dim doc As String
    Dim CLIENTE As String
    Dim ccia As String

    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gCLIENTE
        Get
            Return CLIENTE
        End Get
        Set(value)
            CLIENTE = value
        End Set
    End Property
    Public Property gndoc
        Get
            Return ndoc
        End Get
        Set(value)
            ndoc = value
        End Set
    End Property
    Public Property gdoc
        Get
            Return doc
        End Get
        Set(value)
            doc = value
        End Set
    End Property
    Public Property gvendedor
        Get
            Return vendedor
        End Get
        Set(value)
            vendedor = value
        End Set
    End Property
    Public Property gmes
        Get
            Return mes
        End Get
        Set(value)
            mes = value
        End Set
    End Property
    Public Property gMONTO
        Get
            Return MONTO
        End Get
        Set(value)
            MONTO = value
        End Set
    End Property
    Public Property gsfactu
        Get
            Return sfactu
        End Get
        Set(value)
            sfactu = value
        End Set
    End Property
    Public Property gnfactu
        Get
            Return nfactu
        End Get
        Set(value)
            nfactu = value
        End Set
    End Property
    Public Property gperiodo
        Get
            Return periodo
        End Get
        Set(value)
            periodo = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal sfactu As String, ByVal nfactu As String, ByVal periodo As String, ByVal MONTO As String, ByVal ccia As String,
            ByVal mes As String, ByVal vendedor As String, ByVal doc As String, ByVal ndoc As String, ByVal CLIENTE As String)
        gsfactu = sfactu
        gnfactu = nfactu
        gperiodo = periodo
        gMONTO = MONTO
        gmes = mes
        gvendedor = vendedor
        gdoc = doc
        gndoc = ndoc
        gCLIENTE = CLIENTE
        gccia = ccia
    End Sub
End Class
