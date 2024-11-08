Public Class vacciones

    Dim Accion_realiza As String
    Dim Nombre_fomulario As String
    Dim Codigo_iden As String
    Dim usuario As String
    Dim fyh As DateTime
    Dim maquina As String
    Public Property gAccion_realiza
        Get
            Return Accion_realiza
        End Get
        Set(value)
            Accion_realiza = value
        End Set
    End Property
    Public Property gNombre_fomulario
        Get
            Return Nombre_fomulario
        End Get
        Set(value)
            Nombre_fomulario = value
        End Set
    End Property
    Public Property gCodigo_iden
        Get
            Return Codigo_iden
        End Get
        Set(value)
            Codigo_iden = value
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
    Public Property gfyh
        Get
            Return fyh
        End Get
        Set(value)
            fyh = value
        End Set
    End Property
    Public Property gmaquina
        Get
            Return maquina
        End Get
        Set(value)
            maquina = value
        End Set
    End Property


    Sub New()

    End Sub
    Sub New(ByVal Accion_realiza As String,
        ByVal Nombre_fomulario As String, ByVal Codigo_iden As String, ByVal usuario As String, ByVal fyh As String, ByVal maquina As String
   )
        gAccion_realiza = Accion_realiza
        gNombre_fomulario = Nombre_fomulario
        gCodigo_iden = Codigo_iden
        gusuario = usuario
        gfyh = fyh
        gmaquina = maquina
    End Sub
End Class
