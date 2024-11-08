Public Class vdetdireccion
    Dim codigo As String
    Dim direcion As String
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Public Property gdirecion
        Get
            Return direcion
        End Get
        Set(value)
            direcion = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal codigo As String,
        ByVal direcion As String
   )
        gcodigo = codigo
        gdirecion = direcion

    End Sub
End Class
