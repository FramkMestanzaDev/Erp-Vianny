Public Class vclientevianny
    Dim codigo As String
    Dim nombre As String
    Dim tdoc As String
    Dim ruc As String
    Dim direccion As String
    Dim origen As String
    Dim pais As String
    Dim email_fac As String
    Dim ubigeo As String
    Dim tiper As Integer
    Dim nomf As String
    Dim amaterno As String
    Dim apaterno As String
    Public Property gtiper
        Get
            Return tiper
        End Get
        Set(value)
            tiper = value
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
    Public Property gemail_fac
        Get
            Return email_fac
        End Get
        Set(value)
            email_fac = value
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
    Public Property gorigen
        Get
            Return origen
        End Get
        Set(value)
            origen = value
        End Set
    End Property
    Public Property gdireccion
        Get
            Return direccion
        End Get
        Set(value)
            direccion = value
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
    Public Property gtdoc
        Get
            Return tdoc
        End Get
        Set(value)
            tdoc = value
        End Set
    End Property
    Public Property gnombre
        Get
            Return nombre
        End Get
        Set(value)
            nombre = value
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
    Public Property gnomf
        Get
            Return nomf
        End Get
        Set(value)
            nomf = value
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
    Public Property gapaterno
        Get
            Return apaterno
        End Get
        Set(value)
            apaterno = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal codigo As String,
    ByVal nombre As String,
        ByVal tdoc As String,
        ByVal ruc As String,
        ByVal direccion As String,
        ByVal origen As String,
        ByVal pais As String,
        ByVal email_fac As String,
        ByVal ubigeo As String,
        ByVal tiper As Integer,
            ByVal nomf As String,
        ByVal amaterno As String,
        ByVal apaterno As Integer)
        gcodigo = codigo
        gnombre = nombre
        gtdoc = tdoc
        gruc = ruc
        gdireccion = direccion
        gorigen = origen
        gpais = pais
        gemail_fac = email_fac
        gubigeo = ubigeo
        gtiper = tiper
        gnomf = nomf
        gamaterno = amaterno
        gapaterno = apaterno
    End Sub
End Class
