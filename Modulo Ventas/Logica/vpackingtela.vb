Public Class vpackingtela
    Dim partida As String
    Dim articulo As String
    Dim codigo_trab As String
    Dim nombre_trab As String
    Dim numero_rollos As String
    Dim nrollo As String
    Dim peso As Double
    Dim posicion As String
    Dim almacen As String
    Dim articulo2 As String
    Dim unidad As String
    Dim partida2 As String
    Dim articulo3 As String
    Dim estado As String
    Dim estado1 As String
    Dim codigo_det As Integer
    Dim total As Double
    Dim seleccionado As String
    Dim es As String
    Dim idcabe As String
    Dim codigo2 As Integer
    Dim ROOLO As Integer
    Dim almac As Integer
    Dim calidad As String
    Dim peso_neto As String
    Dim AL As String
    Dim orden As String
    Dim desidad As Double
    Dim ubic_art As String
    Dim guia As String
    Public Property gguia
        Get
            Return guia
        End Get
        Set(value)
            guia = value
        End Set
    End Property
    Public Property gubic_art
        Get
            Return ubic_art
        End Get
        Set(value)
            ubic_art = value
        End Set
    End Property
    Public Property gdesidad
        Get
            Return desidad
        End Get
        Set(value)
            desidad = value
        End Set
    End Property
    Public Property gorden
        Get
            Return orden
        End Get
        Set(value)
            orden = value
        End Set
    End Property
    Public Property gAL
        Get
            Return AL
        End Get
        Set(value)
            AL = value
        End Set
    End Property
    Public Property gpeso_neto
        Get
            Return peso_neto
        End Get
        Set(value)
            peso_neto = value
        End Set
    End Property
    Public Property gcalidad
        Get
            Return calidad
        End Get
        Set(value)
            calidad = value
        End Set
    End Property
    Public Property galmac
        Get
            Return almac
        End Get
        Set(value)
            almac = value
        End Set
    End Property
    Public Property gROOLO
        Get
            Return ROOLO
        End Get
        Set(value)
            ROOLO = value
        End Set
    End Property
    Public Property gcodigo2
        Get
            Return codigo2
        End Get
        Set(value)
            codigo2 = value
        End Set
    End Property
    Public Property gidcabe
        Get
            Return idcabe
        End Get
        Set(value)
            idcabe = value
        End Set
    End Property
    Public Property ges
        Get
            Return es
        End Get
        Set(value)
            es = value
        End Set
    End Property
    Public Property gseleccionado
        Get
            Return seleccionado
        End Get
        Set(value)
            seleccionado = value
        End Set
    End Property
    Public Property gtotal
        Get
            Return total
        End Get
        Set(value)
            total = value
        End Set
    End Property
    Public Property gcodigo_det
        Get
            Return codigo_det
        End Get
        Set(value)
            codigo_det = value
        End Set
    End Property
    Public Property gestado1
        Get
            Return estado1
        End Get
        Set(value)
            estado1 = value
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
    Public Property garticulo3
        Get
            Return articulo3
        End Get
        Set(value)
            articulo3 = value
        End Set
    End Property
    Public Property gpartida2
        Get
            Return partida2
        End Get
        Set(value)
            partida2 = value
        End Set
    End Property
    Public Property gunidad
        Get
            Return unidad
        End Get
        Set(value)
            unidad = value
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
    Public Property garticulo
        Get
            Return articulo
        End Get
        Set(value)
            articulo = value
        End Set
    End Property
    Public Property gcodigo_trab
        Get
            Return codigo_trab
        End Get
        Set(value)
            codigo_trab = value
        End Set
    End Property
    Public Property gnombre_trab
        Get
            Return nombre_trab
        End Get
        Set(value)
            nombre_trab = value
        End Set
    End Property
    Public Property gnumero_rollos
        Get
            Return numero_rollos
        End Get
        Set(value)
            numero_rollos = value
        End Set
    End Property
    Public Property gnrollo
        Get
            Return nrollo
        End Get
        Set(value)
            nrollo = value
        End Set
    End Property
    Public Property gpeso
        Get
            Return peso
        End Get
        Set(value)
            peso = value
        End Set
    End Property
    Public Property gposicion
        Get
            Return posicion
        End Get
        Set(value)
            posicion = value
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
    Public Property garticulo2
        Get
            Return articulo2
        End Get
        Set(value)
            articulo2 = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal partida As String, ByVal total As Double, ByVal seleccionado As String, ByVal es As String, ByVal idcabe As String, ByVal ROOLO As Integer, ByVal calidad As String, ByVal AL As String, ByVal orden As String,
        ByVal articulo As String, ByVal codigo_trab As String, ByVal nombre_trab As String, ByVal numero_rollos As String, ByVal codigo2 As Integer, ByVal almac As String, ByVal peso_neto As String, ByVal desidad As Double, ByVal ubic_art As String,
 ByVal nrollo As String, ByVal peso As Double, ByVal posicion As String, ByVal almacen As String, ByVal articulo2 As String, ByVal unidad As String, ByVal partida2 As String, ByVal articulo3 As String, ByVal estado As String, ByVal estado1 As String, ByVal codigo_det As String)
        gpartida = partida
        garticulo = articulo
        gcodigo_trab = codigo_trab
        gnombre_trab = nombre_trab
        gnumero_rollos = numero_rollos
        gnrollo = nrollo
        gpeso = peso
        gposicion = posicion
        galmacen = almacen
        garticulo2 = articulo2
        gunidad = unidad
        gpartida2 = partida2
        garticulo3 = articulo3
        gestado = estado
        gestado1 = estado1
        gcodigo_det = codigo_det
        gtotal = total
        gseleccionado = seleccionado
        ges = es
        gidcabe = idcabe
        gcodigo2 = codigo2
        gROOLO = ROOLO
        galmac = almac
        gcalidad = calidad
        gpeso_neto = peso_neto
        gAL = AL
        gdesidad = desidad
        gorden = orden
        gubic_art = ubic_art
    End Sub
End Class
