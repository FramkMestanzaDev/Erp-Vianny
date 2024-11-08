Public Class vtejeduria
    Dim tdr_codigo As String
    Dim tdr_descripcion As String
    Dim tdr_usuario As String
    Dim tdr_fupdate As Date
    Dim tdr_pc As String
    Dim tdr_tipodefecto As String
    Dim ccia As String
    Dim rolloini As String
    Dim rollofin As String
    Dim pedidoot As String
    Dim correlativo As String
    Dim ordetejido As String
    Dim fechaini As Date
    Dim fechafin As Date
    Dim tdr_ccia As String
    Dim tsdr_rollo As String
    Dim tdr_defecto As String
    Dim tdr_obs As String
    Dim tdr_cantidad As Integer
    Dim tdr_mtrafectados As Double
    Dim tdr_kgafectados As Double
    Dim maquina As String
    Dim tejedor As String
    Dim po As String
    Dim partida As String
    Public Property gpo
        Get
            Return po
        End Get
        Set(value)
            po = value
        End Set
    End Property
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
        End Set
    End Property
    Public Property gtejedor
        Get
            Return tejedor
        End Get
        Set(value)
            tejedor = value
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
    Public Property gtdr_kgafectados
        Get
            Return tdr_kgafectados
        End Get
        Set(value)
            tdr_kgafectados = value
        End Set
    End Property
    Public Property gtdr_mtrafectados
        Get
            Return tdr_mtrafectados
        End Get
        Set(value)
            tdr_mtrafectados = value
        End Set
    End Property
    Public Property gtdr_cantidad
        Get
            Return tdr_cantidad
        End Get
        Set(value)
            tdr_cantidad = value
        End Set
    End Property
    Public Property gtdr_obs
        Get
            Return tdr_obs
        End Get
        Set(value)
            tdr_obs = value
        End Set
    End Property
    Public Property gtdr_defecto
        Get
            Return tdr_defecto
        End Get
        Set(value)
            tdr_defecto = value
        End Set
    End Property
    Public Property gtsdr_rollo
        Get
            Return tsdr_rollo
        End Get
        Set(value)
            tsdr_rollo = value
        End Set
    End Property
    Public Property gtdr_ccia
        Get
            Return tdr_ccia
        End Get
        Set(value)
            tdr_ccia = value
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
    Public Property gfechaini
        Get
            Return fechaini
        End Get
        Set(value)
            fechaini = value
        End Set
    End Property
    Public Property gordetejido
        Get
            Return ordetejido
        End Get
        Set(value)
            ordetejido = value
        End Set
    End Property
    Public Property gcorrelativo
        Get
            Return correlativo
        End Get
        Set(value)
            correlativo = value
        End Set
    End Property
    Public Property gpedidoot
        Get
            Return pedidoot
        End Get
        Set(value)
            pedidoot = value
        End Set
    End Property
    Public Property grollofin
        Get
            Return rollofin
        End Get
        Set(value)
            rollofin = value
        End Set
    End Property
    Public Property grolloini
        Get
            Return rolloini
        End Get
        Set(value)
            rolloini = value
        End Set
    End Property
    Public Property gccia
        Get
            Return ccia
        End Get
        Set(value)
            ccia = value
        End Set
    End Property
    Public Property gtdr_tipodefecto
        Get
            Return tdr_tipodefecto
        End Get
        Set(value)
            tdr_tipodefecto = value
        End Set
    End Property
    Public Property gtdr_pc
        Get
            Return tdr_pc
        End Get
        Set(value)
            tdr_pc = value
        End Set
    End Property
    Public Property gtdr_fupdate
        Get
            Return tdr_fupdate
        End Get
        Set(value)
            tdr_fupdate = value
        End Set
    End Property
    Public Property gtdr_usuario
        Get
            Return tdr_usuario
        End Get
        Set(value)
            tdr_usuario = value
        End Set
    End Property
    Public Property gtdr_descripcion
        Get
            Return tdr_descripcion
        End Get
        Set(value)
            tdr_descripcion = value
        End Set
    End Property
    Public Property gtdr_codigo
        Get
            Return tdr_codigo
        End Get
        Set(value)
            tdr_codigo = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal tdr_codigo As String, ByVal tdr_descripcion As String, ByVal tdr_usuario As String, ByVal tdr_pc As String, ByVal tdr_tipodefecto As String, ByVal tdr_fupdate As Date, ByVal maquina As String, ByVal tejedor As String,
           ByVal ccia As String, ByVal rolloini As String, ByVal rollofin As String, ByVal pedidoot As String, ByVal correlativo As String, ByVal ordetejido As String, ByVal fechafin As Date, ByVal fechaini As Date, ByVal po As String, ByVal partida As String,
            ByVal tdr_ccia As String, ByVal tsdr_rollo As String, ByVal tdr_defecto As String, ByVal tdr_obs As String, ByVal tdr_cantidad As Integer, ByVal tdr_mtrafectados As Double, ByVal tdr_kgafectados As Double)
        gtdr_codigo = tdr_codigo
        gtdr_descripcion = tdr_descripcion
        gtdr_usuario = tdr_usuario
        gtdr_fupdate = tdr_fupdate
        gtdr_pc = tdr_pc
        gtdr_tipodefecto = tdr_tipodefecto
        gccia = ccia
        grolloini = rolloini
        grollofin = rollofin
        gpedidoot = pedidoot
        gcorrelativo = correlativo
        gordetejido = ordetejido
        gfechafin = fechafin
        gfechaini = fechaini
        gtdr_ccia = tdr_ccia
        gtsdr_rollo = tsdr_rollo
        gtdr_defecto = tdr_defecto
        gtdr_obs = tdr_obs
        gtdr_cantidad = tdr_cantidad
        gtdr_mtrafectados = tdr_mtrafectados
        gtdr_kgafectados = tdr_kgafectados
        gmaquina = maquina
        gtejedor = tejedor
        gpartida = partida
        gpo = po
    End Sub
End Class
