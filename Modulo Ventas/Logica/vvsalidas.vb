Public Class vvsalidas
    Dim mes As String
    Dim ano As String
    Dim almacen As String
    Dim empresa As String
    Dim motivos As String
    Dim codigo As String
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
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
    Public Property gano
        Get
            Return ano
        End Get
        Set(value)
            ano = value
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
    Public Property gempresa
        Get
            Return empresa
        End Get
        Set(value)
            empresa = value
        End Set
    End Property
    Public Property gmotivos
        Get
            Return motivos
        End Get
        Set(value)
            motivos = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal mes As String, ByVal ano As String, ByVal almacen As String, ByVal empresa As String, ByVal motivos As String, ByVal codigo As String)
        gmes = mes
        gano = ano
        gcodigo = codigo
        galmacen = almacen
        gempresa = empresa
        gmotivos = motivos
    End Sub
End Class
