Public Class vreportedoc
    Dim fechaini As Date
    Dim fechafin As Date
    Dim factura As String
    Dim factura2 As String
    Dim estado As String
    Public Property gestado
        Get
            Return estado
        End Get
        Set(value)
            estado = value
        End Set
    End Property
    Public Property gfactura2
        Get
            Return factura2
        End Get
        Set(value)
            factura2 = value
        End Set
    End Property
    Public Property gfactura
        Get
            Return factura
        End Get
        Set(value)
            factura = value
        End Set
    End Property
    Public Property gfechaini
        Get
            Return fechaini
        End Get
        Set(value)
            fechaini = value
        End Set
    End Property
    Public Property gfechafin
        Get
            Return fechafin
        End Get
        Set(value)
            fechafin = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal fechaini As Date,
        ByVal fechafin As Date, ByVal factura As String, ByVal factura2 As String, ByVal estado As String
   )
        gfechaini = fechaini
        gfechafin = fechafin
        gfactura = factura
        gfactura2 = factura2
        gestado = estado
    End Sub
End Class
