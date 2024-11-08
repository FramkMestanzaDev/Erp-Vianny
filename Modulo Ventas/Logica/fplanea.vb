Public Class fplanea
    Dim partida As String
    Dim rollo As String
    Dim fecha As Date
    Dim peso As Double
    Dim ccia As String

    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gpeso
        Get
            Return peso
        End Get
        Set(value)
            peso = value
        End Set
    End Property
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
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
    Public Property grollo
        Get
            Return rollo
        End Get
        Set(value)
            rollo = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal rollo As String, ByVal fecha As Date, ByVal peso As Double, ByVal ccia As String)
        gpartida = partida
        grollo = rollo
        gfecha = fecha
        gpeso = peso
        gccia = ccia
    End Sub
End Class
