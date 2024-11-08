Public Class vsinonimo
    Dim codigo As String
    Dim item As String
    Dim descrip As String

    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Public Property gitem
        Get
            Return item
        End Get
        Set(value)
            item = value
        End Set
    End Property
    Public Property gdescrip
        Get
            Return descrip
        End Get
        Set(value)
            descrip = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal codigo As String, ByVal item As String, ByVal descrip As String)
        gcodigo = codigo
        gitem = item
        gdescrip = descrip
    End Sub
End Class
