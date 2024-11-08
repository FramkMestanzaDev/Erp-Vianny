Public Class vvalidarpartida
    Dim partida As String
    Dim articulo As String
    Dim cantrollos As Double
    Dim kgstejidos As Double
    Dim merma As Double
    Dim rollospesados As Double
    Dim kgsobtenidos As Double
    Dim observacion As String
    Dim empresa As String

    Public Property gempresa
        Get
            Return empresa
        End Get
        Set(value)
            empresa = value
        End Set
    End Property
    Public Property gobservacion
        Get
            Return observacion
        End Get
        Set(value)
            observacion = value
        End Set
    End Property
    Public Property gkgsobtenidos
        Get
            Return kgsobtenidos
        End Get
        Set(value)
            kgsobtenidos = value
        End Set
    End Property
    Public Property grollospesados
        Get
            Return rollospesados
        End Get
        Set(value)
            rollospesados = value
        End Set
    End Property
    Public Property gmerma
        Get
            Return merma
        End Get
        Set(value)
            merma = value
        End Set
    End Property
    Public Property gkgstejidos
        Get
            Return kgstejidos
        End Get
        Set(value)
            kgstejidos = value
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
    Public Property garticulo
        Get
            Return articulo
        End Get
        Set(value)
            articulo = value
        End Set
    End Property
    Public Property gcantrollos
        Get
            Return cantrollos
        End Get
        Set(value)
            cantrollos = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal articulo As String, ByVal kgstejidos As Double, ByVal merma As Double, ByVal rollospesados As Double, ByVal kgsobtenidos As Double, ByVal cantrollos As Double, ByVal empresa As String,
        ByVal observacion As String)
        gpartida = partida
        garticulo = articulo
        gcantrollos = cantrollos
        gkgstejidos = kgstejidos
        gmerma = merma
        grollospesados = rollospesados
        gkgsobtenidos = kgsobtenidos
        gobservacion = observacion
        gempresa = empresa
    End Sub
End Class
