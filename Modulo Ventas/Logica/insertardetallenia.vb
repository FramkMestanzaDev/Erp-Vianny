Public Class insertardetallenia
    Dim ccia As String
    Dim ncom As String
    Dim item As String
    Dim linea As String
    Dim cantidad As Double
    Dim cantidad1 As Double
    Dim und As String
    Dim op As String
    Dim partida As String
    Dim almacen As String
    Dim umpresentacion As String
    Dim otejeduria As String
    Dim oc As String
    Dim ano As String
    Dim lote As String
    Dim cenvio As Double
    Dim clasif As String

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
    Public Property gcenvio
        Get
            Return cenvio
        End Get
        Set(value)
            cenvio = value
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
    Public Property goc
        Get
            Return oc
        End Get
        Set(value)
            oc = value
        End Set
    End Property
    Public Property gumpresentacion
        Get
            Return umpresentacion
        End Get
        Set(value)
            umpresentacion = value
        End Set
    End Property
    Public Property gotejeduria
        Get
            Return otejeduria
        End Get
        Set(value)
            otejeduria = value
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
    Public Property gpartida
        Get
            Return partida
        End Get
        Set(value)
            partida = value
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
    Public Property gcantidad1
        Get
            Return cantidad1
        End Get
        Set(value)
            cantidad1 = value
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
    Public Property galmacen
        Get
            Return almacen
        End Get
        Set(value)
            almacen = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal ncom As String, ByVal item As String, ByVal lote As String, ByVal cenvio As Double, ByVal clasif As String,
            ByVal linea As String, ByVal cantidad As Double, ByVal cantidad1 As Double, ByVal oc As String, ByVal ano As String, ByVal ccia As String,
            ByVal und As String, ByVal op As String, ByVal almacen As String, ByVal umpresentacion As String, ByVal otejeduria As String)
        gncom = ncom
        gitem = item
        glinea = linea
        gcantidad = cantidad
        gcantidad1 = cantidad1
        gund = und
        gop = op
        gpartida = partida
        galmacen = almacen
        gumpresentacion = umpresentacion
        gotejeduria = otejeduria
        goc = oc
        gano = ano
        glote = lote
        gcenvio = cenvio
        gccia = ccia
        gclasif = clasif
    End Sub
End Class
