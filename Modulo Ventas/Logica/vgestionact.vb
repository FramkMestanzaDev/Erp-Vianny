Public Class vgestionact
    Dim id_ga As String
    Dim fecha_g As Date
    Dim vededor_co As String
    Dim vendedor_nom As String
    Dim cliente_ruc As String
    Dim cliente_rs As String
    Dim prox_visita As Date
    Dim contacto As String
    Dim estado As String
    Dim descripcion As String
    Dim ubicacion As String
    Dim fini As Date
    Dim ffin As Date
    Dim CORREO As String
    Dim CELULAR As String
    Dim MONEDA As String
    Dim MONTO As Double
    Dim CONTACTARON As String
    Public Property gMONEDA
        Get
            Return MONEDA
        End Get
        Set(value)
            MONEDA = value
        End Set
    End Property
    Public Property gMONTO
        Get
            Return MONTO
        End Get
        Set(value)
            MONTO = value
        End Set
    End Property
    Public Property gCORREO
        Get
            Return CORREO
        End Get
        Set(value)
            CORREO = value
        End Set
    End Property
    Public Property gCELULAR
        Get
            Return CELULAR
        End Get
        Set(value)
            CELULAR = value
        End Set
    End Property
    Public Property gfini
        Get
            Return fini
        End Get
        Set(value)
            fini = value
        End Set
    End Property
    Public Property gffin
        Get
            Return ffin
        End Get
        Set(value)
            ffin = value
        End Set
    End Property
    Public Property gubicacion
        Get
            Return ubicacion
        End Get
        Set(value)
            ubicacion = value
        End Set
    End Property
    Public Property gdescripcion
        Get
            Return descripcion
        End Get
        Set(value)
            descripcion = value
        End Set
    End Property
    Public Property gestado
        Get
            Return estado
        End Get
        Set(value)
            estado = value
        End Set
    End Property
    Public Property gcontacto
        Get
            Return contacto
        End Get
        Set(value)
            contacto = value
        End Set
    End Property
    Public Property gprox_visita
        Get
            Return prox_visita
        End Get
        Set(value)
            prox_visita = value
        End Set
    End Property
    Public Property gcliente_rs
        Get
            Return cliente_rs
        End Get
        Set(value)
            cliente_rs = value
        End Set
    End Property
    Public Property gcliente_ruc
        Get
            Return cliente_ruc
        End Get
        Set(value)
            cliente_ruc = value
        End Set
    End Property
    Public Property gvendedor_nom
        Get
            Return vendedor_nom
        End Get
        Set(value)
            vendedor_nom = value
        End Set
    End Property
    Public Property gid_ga
        Get
            Return id_ga
        End Get
        Set(value)
            id_ga = value
        End Set
    End Property
    Public Property gfecha_g
        Get
            Return fecha_g
        End Get
        Set(value)
            fecha_g = value
        End Set
    End Property
    Public Property gCONTACTARON
        Get
            Return CONTACTARON
        End Get
        Set(value)
            CONTACTARON = value
        End Set
    End Property
    Public Property gvededor_co
        Get
            Return vededor_co
        End Get
        Set(value)
            vededor_co = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal id_ga As String, ByVal fecha_g As Date, ByVal vededor_co As String, ByVal vendedor_nom As String, ByVal cliente_ruc As String,
         ByVal cliente_rs As String, ByVal prox_visita As Date, ByVal contacto As String, ByVal estado As String, ByVal descripcion As String,
            ByVal ubicacion As String, ByVal ffin As Date, ByVal fini As Date, ByVal CORREO As String, ByVal CELULAR As String, ByVal MONEDA As String, ByVal MONTO As Double, ByVal CONTACTARON As String)
        gid_ga = id_ga
        gfecha_g = fecha_g
        gvededor_co = vededor_co
        gvendedor_nom = vendedor_nom
        gcliente_ruc = cliente_ruc
        gcliente_rs = cliente_rs
        gprox_visita = prox_visita
        gcontacto = contacto
        gestado = estado
        gdescripcion = descripcion
        gubicacion = ubicacion
        gffin = ffin
        gfini = fini
        gCORREO = CORREO
        gCELULAR = CELULAR
        gMONEDA = MONEDA
        gMONTO = MONTO
        gCONTACTARON = CONTACTARON
    End Sub
End Class
