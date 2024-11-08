Public Class vestadocuenta
    Dim ruc As String
    Dim cperiodo As String
    Dim FECHA As Date
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
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
    Public Property gcperiodo
        Get
            Return cperiodo
        End Get
        Set(value)
            cperiodo = value
        End Set
    End Property
    Public Property gFECHA
        Get
            Return FECHA
        End Get
        Set(value)
            FECHA = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ruc As String, ByVal cperiodo As String, ByVal FECHA As Date, ByVal ccia As String)
        gruc = ruc
        gcperiodo = cperiodo
        gFECHA = FECHA
        gccia = ccia
    End Sub
End Class
