Public Class vnsadetalle
    Dim ncom As String
    Dim item As String
    Dim linea As String
    Dim cantidad As Double
    Dim und As String
    Dim op As String
    Dim partida As String
    Dim rollo1 As Double
    Dim almacen As String
    Dim unidad_medidad As String
    Dim ordtejeduria As String
    Dim ano As String
    Dim lote As String
    Dim cantenvio As Double
    Dim ccia As String
    Dim clasif As String
    Dim ocompra As String

    Public Property gclasif
        Get
            Return clasif
        End Get
        Set(value)
            clasif = value
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
    Public Property gcantenvio
        Get
            Return cantenvio
        End Get
        Set(value)
            cantenvio = value
        End Set
    End Property
    Public Property glote
        Get
            Return lote
        End Get
        Set(value)
            lote = value
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
    Public Property gordtejeduria
        Get
            Return ordtejeduria
        End Get
        Set(value)
            ordtejeduria = value
        End Set
    End Property
    Public Property gunidad_medidad
        Get
            Return unidad_medidad
        End Get
        Set(value)
            unidad_medidad = value
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
    Public Property gitem
        Get
            Return item
        End Get
        Set(value)
            item = value
        End Set
    End Property
    Public Property glinea
        Get
            Return linea
        End Get
        Set(value)
            linea = value
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
    Public Property gund
        Get
            Return und
        End Get
        Set(value)
            und = value
        End Set
    End Property
    Public Property gop
        Get
            Return op
        End Get
        Set(value)
            op = value
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
    Public Property grollo1
        Get
            Return rollo1
        End Get
        Set(value)
            rollo1 = value
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

    Public Property gOcompra As String
        Get
            Return ocompra
        End Get
        Set(value As String)
            ocompra = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(
            ByVal ncom As String, ByVal cantenvio As Double,
    ByVal item As String,
        ByVal linea As String,
        ByVal cantidad As Double,
        ByVal und As String, ByVal clasif As String,
        ByVal op As String, ByVal lote As String,
        ByVal partida As String, ByVal ano As String, ByVal ccia As String, ByVal ocompra As String,
        ByVal rollo1 As Double, ByVal almacen As String, ByVal unidad_medidad As String, ByVal ordtejeduria As String
        )
        gclasif = clasif
        gncom = ncom
        gitem = item
        glinea = linea
        gcantidad = cantidad
        gund = und
        gop = op
        gpartida = partida
        grollo1 = rollo1
        galmacen = almacen
        gunidad_medidad = unidad_medidad
        gordtejeduria = ordtejeduria
        gano = ano
        glote = lote
        gcantenvio = cantenvio
        gccia = ccia
        gOcompra = ocompra
    End Sub
End Class
