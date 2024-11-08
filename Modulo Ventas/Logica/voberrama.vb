Public Class voberrama

    Dim PROGRAMA As String
    Dim FECHA As Date
    Dim HORA As String
    Dim MOTIVO As String
    Dim COMENTARIO As String
    Public Property gCOMENTARIO
        Get
            Return COMENTARIO
        End Get
        Set(value)
            COMENTARIO = value
        End Set
    End Property
    Public Property gMOTIVO
        Get
            Return MOTIVO
        End Get
        Set(value)
            MOTIVO = value
        End Set
    End Property
    Public Property gHORA
        Get
            Return HORA
        End Get
        Set(value)
            HORA = value
        End Set
    End Property
    Public Property gPROGRAMA
        Get
            Return PROGRAMA
        End Get
        Set(value)
            PROGRAMA = value
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
    Sub New(ByVal fecha As Date, ByVal hora As String, ByVal motivo As String, ByVal comentario As String,
        ByVal PROGRAMA As String)
        gPROGRAMA = PROGRAMA
        gFECHA = fecha
        gHORA = hora
        gMOTIVO = motivo
        gCOMENTARIO = comentario

    End Sub
End Class
