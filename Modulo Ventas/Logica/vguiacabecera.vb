Public Class vguiacabecera
    Dim sfactu As String
    Dim nfactu As String
    Dim ruc As String
    Dim rsocial As String
    Dim fcom As Date
    Dim fcom1 As Date
    Dim direccion As String
    Dim placa As String
    Dim direccion_partida As String
    Dim NOTASALIDA As String
    Dim chofer As String
    Dim serie As String
    Dim correlativo As String
    Dim almacen As String
    Dim licencia As String
    Dim ruc_transpo As String
    Dim direc_transport As String
    Dim ncom As String
    Dim ruc3 As String
    Dim tip_documento As String
    Dim motivo As String
    Dim destino As String
    Dim usuario As String
    Dim fase As String
    Dim mes As String
    Dim ccia As String
    Dim CodArtIni As String
    Dim FiltroDescrip As String
    Dim Modal As String
    Dim trpo As String
    Dim fini As Date
    Dim ffin As Date
    Dim ano As String
    Dim codigo As String
    Dim partida As String
    Dim cantidad As Double
    Dim glosa As String
    Dim ordens_3 As String
    Dim subfase_3 As String
    Dim situacion As String
    Public Property gfcom1
        Get
            Return fcom1
        End Get
        Set(value)
            fcom1 = value
        End Set
    End Property
    Public Property gsituacion
        Get
            Return situacion
        End Get
        Set(value)
            situacion = value
        End Set
    End Property
    Public Property gsubfase_3
        Get
            Return subfase_3
        End Get
        Set(value)
            subfase_3 = value
        End Set
    End Property
    Public Property gordens_3
        Get
            Return ordens_3
        End Get
        Set(value)
            ordens_3 = value
        End Set
    End Property
    Public Property gglosa
        Get
            Return glosa
        End Get
        Set(value)
            glosa = value
        End Set
    End Property
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
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
    Public Property gcantidad
        Get
            Return cantidad
        End Get
        Set(value)
            cantidad = value
        End Set
    End Property
    Public Property gano
        Get
            Return ano
        End Get
        Set(value)
            ano = value
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
    Public Property gtrpo
        Get
            Return trpo
        End Get
        Set(value)
            trpo = value
        End Set
    End Property
    Public Property gModal
        Get
            Return Modal
        End Get
        Set(value)
            Modal = value
        End Set
    End Property
    Public Property gFiltroDescrip
        Get
            Return FiltroDescrip
        End Get
        Set(value)
            FiltroDescrip = value
        End Set
    End Property
    Public Property gCodArtIni
        Get
            Return CodArtIni
        End Get
        Set(value)
            CodArtIni = value
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
    Public Property gmes
        Get
            Return mes
        End Get
        Set(value)
            mes = value
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
    Public Property gfase
        Get
            Return fase
        End Get
        Set(value)
            fase = value
        End Set
    End Property
    Public Property gmotivo
        Get
            Return motivo
        End Get
        Set(value)
            motivo = value
        End Set
    End Property
    Public Property gdestino
        Get
            Return destino
        End Get
        Set(value)
            destino = value
        End Set
    End Property
    Public Property gtip_documento
        Get
            Return tip_documento
        End Get
        Set(value)
            tip_documento = value
        End Set
    End Property
    Public Property gruc3
        Get
            Return ruc3
        End Get
        Set(value)
            ruc3 = value
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
    Public Property gserie
        Get
            Return serie
        End Get
        Set(value)
            serie = value
        End Set
    End Property
    Public Property gsfactu
        Get
            Return sfactu
        End Get
        Set(value)
            sfactu = value
        End Set
    End Property
    Public Property gnfactu
        Get
            Return nfactu
        End Get
        Set(value)
            nfactu = value
        End Set
    End Property
    Public Property gruc
        Get
            Return ruc
        End Get
        Set(value)
            ruc = value
        End Set
    End Property
    Public Property grsocial
        Get
            Return rsocial
        End Get
        Set(value)
            rsocial = value
        End Set
    End Property
    Public Property gfcom
        Get
            Return fcom
        End Get
        Set(value)
            fcom = value
        End Set
    End Property
    Public Property gdireccion
        Get
            Return direccion
        End Get
        Set(value)
            direccion = value
        End Set
    End Property
    Public Property gplaca
        Get
            Return placa
        End Get
        Set(value)
            placa = value
        End Set
    End Property
    Public Property gdireccion_partida
        Get
            Return direccion_partida
        End Get
        Set(value)
            direccion_partida = value
        End Set
    End Property
    Public Property gNOTASALIDA
        Get
            Return NOTASALIDA
        End Get
        Set(value)
            NOTASALIDA = value
        End Set
    End Property
    Public Property gchofer
        Get
            Return chofer
        End Get
        Set(value)
            chofer = value
        End Set
    End Property
    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
        End Set
    End Property

    Public Property gruc_transpo
        Get
            Return ruc_transpo
        End Get
        Set(value)
            ruc_transpo = value
        End Set
    End Property
    Public Property glicencia
        Get
            Return licencia
        End Get
        Set(value)
            licencia = value
        End Set
    End Property
    Public Property gdirec_transport
        Get
            Return direc_transport
        End Get
        Set(value)
            direc_transport = value
        End Set
    End Property

    Public Property gncom
        Get
            Return ncom
        End Get
        Set(value)
            ncom = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal mes As String, ByVal sfactu As String, ByVal nfactu As String, ByVal ruc As String, ByVal rsocial As String, ByVal glosa As String, ByVal situacion As String,
            ByVal fcom As Date, ByVal direccion As String, ByVal placa As String, ByVal direccion_partida As String, ByVal NOTASALIDA As String,
            ByVal chofer As String, ByVal almacen As String, ByVal licencia As String, ByVal ruc_transpo As String, ByVal direc_transport As String, ByVal ano As String,
            ByVal ncom As String, ByVal ruc3 As String, ByVal tip_documento As String, ByVal motivo As String, ByVal destino As String, ByVal ffin As Date, ByVal fini As Date,
            ByVal usuario As String, ByVal fase As String, ByVal ccia As String, ByVal CodArtIni As String, ByVal FiltroDescrip As String, ByVal Modal As String, ByVal trpo As String,
            ByVal codigo As String, ByVal partida As String, ByVal cantidad As Double, ByVal ordens_3 As String, ByVal subfase_3 As String, ByVal fcom1 As Date)
        gsfactu = sfactu
        gnfactu = nfactu
        gruc = ruc
        grsocial = rsocial
        gfcom = fcom
        gfcom1 = fcom1
        gdireccion = direccion
        gplaca = placa
        gdireccion_partida = direccion_partida
        gNOTASALIDA = NOTASALIDA
        gchofer = chofer
        galmacen = almacen
        glicencia = licencia
        gruc_transpo = ruc_transpo
        gdirec_transport = direc_transport
        gncom = ncom
        gruc3 = ruc3
        gtip_documento = tip_documento
        gdestino = destino
        gmotivo = motivo
        gfase = fase
        gusuario = usuario
        gmes = mes
        gccia = ccia
        gCodArtIni = CodArtIni
        gFiltroDescrip = FiltroDescrip
        gModal = Modal
        gtrpo = trpo
        gffin = ffin
        gfini = fini
        gano = ano
        gcodigo = codigo
        gpartida = partida
        gcantidad = cantidad
        gglosa = glosa
        gordens_3 = ordens_3
        gsubfase_3 = subfase_3
        gsituacion = situacion
    End Sub
End Class
