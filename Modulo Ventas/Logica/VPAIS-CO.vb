Public Class VPAIS_CO

    Dim ruc As String
    Dim orden As String
    Public Property gruc
        Get
            Return ruc
        End Get
        Set(value)
            ruc = value
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
    Sub New()

    End Sub
    Sub New(ByVal ruc As String, ByVal orden As String
   )
        gruc = ruc
        gorden = orden
    End Sub
End Class
