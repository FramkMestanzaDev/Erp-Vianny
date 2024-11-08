Public Class vproducot
    Dim cia As String
    Dim t_pedido As String
    Dim codigo As String
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Public Property gcia
        Get
            Return cia
        End Get
        Set(value)
            cia = value
        End Set
    End Property
    Public Property gt_pedido
        Get
            Return t_pedido
        End Get
        Set(value)
            t_pedido = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal cia As String,
        ByVal t_pedido As String, ByVal codigo As String)
        gcia = cia
        gt_pedido = t_pedido
        gcodigo = codigo
    End Sub
End Class
