Public Class vingre_tranportistas
    Dim codigo As String
    Dim chofer As String
    Dim marca As String
    Dim placa As String
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Public Property gchofer
        Get
            Return chofer
        End Get
        Set(value)
            chofer = value
        End Set
    End Property
    Public Property gmarca
        Get
            Return marca
        End Get
        Set(value)
            marca = value
        End Set
    End Property
    Public Property gplaca
        Get
            Return placa
        End Get
        Set(value)
            placa = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal codigo As String, ByVal chofer As String, ByVal marca As String, ByVal placa As String)
        gcodigo = codigo
        gchofer = chofer
        gmarca = marca
        gplaca = placa

    End Sub
End Class
