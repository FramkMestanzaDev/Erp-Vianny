Public Class vpartida
    Dim partida As String
    Dim codigo As String
    Dim almacen As String
    Dim rollo As String
    Public Property grollo
        Get
            Return rollo
        End Get
        Set(value)
            rollo = value
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
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
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
    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal codigo As String, ByVal almacen As String, ByVal rollo As String)
        gpartida = partida
        gcodigo = codigo
        galmacen = almacen
        grollo = rollo
    End Sub
End Class
