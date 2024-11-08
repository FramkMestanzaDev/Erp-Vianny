Public Class vpackin_tela_det
    Dim rollo As String
    Dim cantidad As Double
    Dim ubicacion As String
    Dim almacen As String
    Dim id_packing As String
    Dim codigo_pro As String
    Dim partida As String
    Dim ESTADO As String
    Public Property gestado
        Get
            Return ESTADO
        End Get
        Set(value)
            ESTADO = value
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
    Public Property gcantidad
        Get
            Return cantidad
        End Get
        Set(value)
            cantidad = value
        End Set
    End Property
    Public Property gubicacion
        Get
            Return ubicacion
        End Get
        Set(value)
            ubicacion = value
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
    Public Property gid_packing
        Get
            Return id_packing
        End Get
        Set(value)
            id_packing = value
        End Set
    End Property
    Public Property gcodigo_pro
        Get
            Return codigo_pro
        End Get
        Set(value)
            codigo_pro = value
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
    Sub New()

    End Sub
    Sub New(ByVal rollo As String, ByVal cantidad As Double, ByVal ubicacion As String, ByVal estado As String,
            ByVal codigo_pro As String, ByVal almacen As String, ByVal id_packing As String,
            ByVal partida As String)

        grollo = rollo
        gcantidad = cantidad
        gubicacion = ubicacion
        galmacen = almacen
        gid_packing = id_packing
        gcodigo_pro = codigo_pro
        gpartida = partida
        gestado = estado
    End Sub
End Class
