Public Class fguia2
    Dim sfactu As String
    Dim nfactu As String
    Dim ccia As String
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gsfactu
        Get
            Return sfactu
        End Get
        Set(value)
            sfactu = value
        End Set
    End Property
    Public Property gnfactu
        Get
            Return nfactu
        End Get
        Set(value)
            nfactu = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal sfactu As String, ByVal ccia As String,
             ByVal nfactu As String
   )
        gccia = ccia
        gsfactu = sfactu
        gnfactu = nfactu
    End Sub
End Class
