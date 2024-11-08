Public Class vlog
    Dim modulo As String
    Dim cuser As String
    Dim fecha As DateTime
    Dim pc As String
    Dim accion As String
    Dim detalle As String
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gmodulo
        Get
            Return modulo
        End Get
        Set(value)
            modulo = value
        End Set
    End Property
    Public Property gcuser
        Get
            Return cuser
        End Get
        Set(value)
            cuser = value
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
    Public Property gpc
        Get
            Return pc
        End Get
        Set(value)
            pc = value
        End Set
    End Property
    Public Property gaccion
        Get
            Return accion
        End Get
        Set(value)
            accion = value
        End Set
    End Property
    Public Property gdetalle
        Get
            Return detalle
        End Get
        Set(value)
            detalle = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal modulo As String, ByVal cuser As String, ByVal fecha As DateTime, ByVal pc As String, ByVal accion As String, ByVal detalle As String, ByVal ccia As String)

        gmodulo = modulo
        gcuser = cuser
        gfecha = fecha
        gpc = pc
        gaccion = accion
        gdetalle = detalle
        gccia = ccia

    End Sub
End Class
