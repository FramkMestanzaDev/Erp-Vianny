Public Class vanalisis
    Dim fini As Date
    Dim ffin As Date
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gfini
        Get
            Return fini
        End Get
        Set(value)
            fini = value
        End Set
    End Property
    Public Property gffin
        Get
            Return ffin
        End Get
        Set(value)
            ffin = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal fini As Date, ByVal ccia As String,
        ByVal ffin As Date)
        gfini = fini
        gffin = ffin
        gccia = ccia
    End Sub
End Class
