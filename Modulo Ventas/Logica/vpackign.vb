Public Class vpackign
    Dim almacen As String
    Dim estado As String
    Dim ID As String
    Dim partida As String
    Dim selec As String
    Dim guia As String
    Dim calidad As String
    Dim Npedido As String
    Public Property gNpedido
        Get
            Return Npedido
        End Get
        Set(value)
            Npedido = value
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
    Public Property gselec
        Get
            Return selec
        End Get
        Set(value)
            selec = value
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
    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
        End Set
    End Property
    Public Property gID
        Get
            Return ID
        End Get
        Set(value)
            ID = value
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

    Public Property gCalidad As String
        Get
            Return calidad
        End Get
        Set(value As String)
            calidad = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal estado As String, ByVal Npedido As String, ByVal almacen As String, ByVal ID As String, ByVal partida As String, ByVal selec As String, ByVal guia As String, ByVal Calidad As String)
        galmacen = almacen
        gestado = estado
        gID = ID
        gpartida = partida
        gselec = selec
        gguia = guia
        gCalidad = Calidad
        gNpedido = Npedido
    End Sub
End Class
