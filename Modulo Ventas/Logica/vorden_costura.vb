Public Class vorden_costura
    Dim op As String
    Dim cantidad As Double
    Dim orden As String
    Dim orden_c As String
    Dim op_oc As String
    Dim ocorte As String
    Dim prendasc As Integer
    Dim adelanto As Integer
    Dim ruc As String
    Dim fecha As Date
    Public Property gorden_c
        Get
            Return orden_c
        End Get
        Set(value)
            orden_c = value
        End Set
    End Property
    Public Property gop_oc
        Get
            Return op_oc
        End Get
        Set(value)
            op_oc = value
        End Set
    End Property
    Public Property gocorte
        Get
            Return ocorte
        End Get
        Set(value)
            ocorte = value
        End Set
    End Property
    Public Property gprendasc
        Get
            Return prendasc
        End Get
        Set(value)
            prendasc = value
        End Set
    End Property
    Public Property gadelanto
        Get
            Return adelanto
        End Get
        Set(value)
            adelanto = value
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
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Public Property gop
        Get
            Return op
        End Get
        Set(value)
            op = value
        End Set
    End Property
    Public Property gorden
        Get
            Return orden
        End Get
        Set(value)
            orden = value
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
    Sub New()

    End Sub
    Sub New(ByVal op As String, ByVal cantidad As Double, ByVal orden As String, ByVal orden_c As String,
            ByVal op_oc As String, ByVal ocorte As String, ByVal prendasc As Integer, ByVal adelanto As Integer,
            ByVal ruc As String, ByVal fecha As Date)
        gop = op
        gcantidad = cantidad
        gorden = orden
        gorden_c = orden_c
        gop_oc = op_oc
        gocorte = ocorte
        gprendasc = prendasc
        gadelanto = adelanto
        gruc = ruc
        gfecha = fecha
    End Sub
End Class
