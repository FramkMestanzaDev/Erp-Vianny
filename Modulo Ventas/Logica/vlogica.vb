Public Class vlogica
    Dim CCia As String
    Dim Familia As String
    Dim SubFamilia As String
    Dim Origen As String
    Dim Color As String
    Public Property gCCia
        Get
            Return CCia
        End Get
        Set(value)
            CCia = value
        End Set
    End Property
    Public Property gFamilia
        Get
            Return Familia
        End Get
        Set(value)
            Familia = value
        End Set
    End Property
    Public Property gSubFamilia
        Get
            Return SubFamilia
        End Get
        Set(value)
            SubFamilia = value
        End Set
    End Property
    Public Property gOrigen
        Get
            Return Origen
        End Get
        Set(value)
            Origen = value
        End Set
    End Property
    Public Property gColor
        Get
            Return Color
        End Get
        Set(value)
            Color = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal CCia As String, ByVal Familia As String, ByVal SubFamilia As String, ByVal Origen As String, ByVal Color As String)
        gCCia = CCia
        gFamilia = Familia
        gSubFamilia = SubFamilia
        gOrigen = Origen
        gColor = Color
    End Sub
End Class
