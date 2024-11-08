Public Class vmatp041f
    Dim ccia_41A As String
    Dim ncom_41a As String
    Dim itm_41a As String
    Dim hilo_41a As String
    Dim cant_41a As Double
    Dim lote_41a As String
    Dim prvhil_41a As String
    Public Property gprvhil_41a
        Get
            Return prvhil_41a
        End Get
        Set(value)
            prvhil_41a = value
        End Set
    End Property
    Public Property glote_41a
        Get
            Return lote_41a
        End Get
        Set(value)
            lote_41a = value
        End Set
    End Property
    Public Property gcant_41a
        Get
            Return cant_41a
        End Get
        Set(value)
            cant_41a = value
        End Set
    End Property
    Public Property ghilo_41a
        Get
            Return hilo_41a
        End Get
        Set(value)
            hilo_41a = value
        End Set
    End Property
    Public Property gitm_41a
        Get
            Return itm_41a
        End Get
        Set(value)
            itm_41a = value
        End Set
    End Property
    Public Property gncom_41a
        Get
            Return ncom_41a
        End Get
        Set(value)
            ncom_41a = value
        End Set
    End Property
    Public Property gccia_41A
        Get
            Return ccia_41A
        End Get
        Set(value)
            ccia_41A = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal ccia_41A As String,
    ByVal ncom_41a As String,
        ByVal itm_41a As String,
        ByVal hilo_41a As String,
        ByVal cant_41a As Double,
        ByVal lote_41a As String,
        ByVal prvhil_41a As String)
        gccia_41A = ccia_41A
        gncom_41a = ncom_41a
        gitm_41a = itm_41a
        ghilo_41a = hilo_41a
        gcant_41a = cant_41a
        glote_41a = lote_41a
        gprvhil_41a = prvhil_41a
    End Sub
End Class
