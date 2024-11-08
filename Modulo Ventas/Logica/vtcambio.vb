Public Class vtcambio
    Dim fecha As Date
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal fecha As Date)
        gfecha = fecha

    End Sub
End Class
