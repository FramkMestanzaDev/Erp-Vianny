Public Class vusuario
    Dim id_usuario As String
    Dim nombre As String
    Dim usuario As String
    Dim clave As String
    Dim grupo As String
    Dim correo As String


    Public Property gid_usuario
        Get
            Return id_usuario
        End Get
        Set(value)
            id_usuario = value
        End Set
    End Property
    Public Property gnombre
        Get
            Return nombre
        End Get
        Set(value)
            nombre = value
        End Set
    End Property
    Public Property gusuario
        Get
            Return usuario
        End Get
        Set(value)
            usuario = value
        End Set
    End Property
    Public Property gclave
        Get
            Return clave
        End Get
        Set(value)
            clave = value
        End Set
    End Property
    Public Property ggrupo
        Get
            Return grupo
        End Get
        Set(value)
            grupo = value
        End Set
    End Property
    Public Property gcorreo
        Get
            Return correo
        End Get
        Set(value)
            correo = value
        End Set
    End Property


    Sub New()

    End Sub
    Sub New(ByVal id_usuario As String,
        ByVal nombre As String, ByVal clave As String, ByVal usuario As String, ByVal grupo As String, ByVal correo As String
   )
        gid_usuario = id_usuario
        gnombre = nombre
        gclave = clave
        gusuario = usuario
        ggrupo = grupo
        gcorreo = correo
    End Sub
End Class
