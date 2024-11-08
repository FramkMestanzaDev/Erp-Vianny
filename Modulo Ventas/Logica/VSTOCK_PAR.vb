Public Class VSTOCK_PAR
    Dim CCIA As String
    Dim ALMACEN As String
    Dim MODAL As String
    Dim po As String
    Dim ano As String
    Dim codigo As String
    Dim rubro As String

    Dim ocompra As String

    Dim cuenta As String
    Dim registro As String
    Dim periodo As String
    Dim dh As String
    Dim ccom As String
    Public Property gccom
        Get
            Return ccom
        End Get
        Set(value)
            ccom = value
        End Set
    End Property
    Public Property gdh
        Get
            Return dh
        End Get
        Set(value)
            dh = value
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
    Public Property gregistro
        Get
            Return registro
        End Get
        Set(value)
            registro = value
        End Set
    End Property
    Public Property gcuenta
        Get
            Return cuenta
        End Get
        Set(value)
            cuenta = value
        End Set
    End Property
    Public Property gocompra
        Get
            Return ocompra
        End Get
        Set(value)
            ocompra = value
        End Set
    End Property
    Public Property grubro
        Get
            Return rubro
        End Get
        Set(value)
            rubro = value
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
    Public Property gano
        Get
            Return ano
        End Get
        Set(value)
            ano = value
        End Set
    End Property
    Public Property gpo
        Get
            Return po
        End Get
        Set(value)
            po = value
        End Set
    End Property
    Public Property gCCIA
        Get
            Return CCIA
        End Get
        Set(value)
            CCIA = value
        End Set
    End Property
    Public Property gALMACEN
        Get
            Return ALMACEN
        End Get
        Set(value)
            ALMACEN = value
        End Set
    End Property
    Public Property gMODAL
        Get
            Return MODAL
        End Get
        Set(value)
            MODAL = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal ALMACEN As String, ByVal MODAL As String, ByVal po As String, ByVal ano As String, ByVal codigo As String, ByVal rubro As String, ByVal ocompra As String,
             ByVal cuenta As String, ByVal registro As String, ByVal periodo As String, ByVal dh As String, ByVal ccom As String)
        gCCIA = CCIA
        gALMACEN = ALMACEN
        gMODAL = MODAL
        gpo = po
        gano = ano
        gcodigo = codigo
        grubro = rubro

        gocompra = ocompra

        gcuenta = cuenta
        gregistro = registro
        gperiodo = periodo
        gdh = dh
        gccom = ccom
    End Sub
End Class
