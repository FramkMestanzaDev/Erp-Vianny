Public Class vlibros
    Dim ccia As String
    Dim CPERIODO As String
    Dim fecha_ini As Date
    Dim fecha_FIN As Date
    Dim mes As String
    Public Property gmes
        Get
            Return mes
        End Get
        Set(value)
            mes = value
        End Set
    End Property
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gCPERIODO
        Get
            Return CPERIODO
        End Get
        Set(value)
            CPERIODO = value
        End Set
    End Property
    Public Property gfecha_ini
        Get
            Return fecha_ini
        End Get
        Set(value)
            fecha_ini = value
        End Set
    End Property
    Public Property gfecha_FIN
        Get
            Return fecha_FIN
        End Get
        Set(value)
            fecha_FIN = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ccia As String, ByVal CPERIODO As String, ByVal fecha_ini As Date, ByVal fecha_FIN As Date, ByVal mes As String)
        gccia = ccia
        gCPERIODO = CPERIODO
        gfecha_ini = fecha_ini
        gfecha_FIN = fecha_FIN
        gmes = mes
    End Sub
End Class
