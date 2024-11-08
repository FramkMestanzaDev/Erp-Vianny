Public Class vguiadetalle
    Dim sfactu As String
    Dim nfactu As String
    Dim items As String
    Dim linea As String
    Dim cantidad As Double
    Dim canencio As Double
    Dim ordens As String
    Dim parti As String
    Dim rollo As Double
    Dim unidad_medida As String
    Dim unidad_medida2 As String
    Dim lote As String
    Dim ccia As String
    Dim ordtejeduria As String
    Dim obser1_3a As String
    Dim obser2_3a As String
    Dim obser3_3a As String
    Dim obser4_3a As String
    Dim obser5_3a As String
    Dim clasif As String
    Public Property gclasif
        Get
            Return clasif
        End Get
        Set(value)
            clasif = value
        End Set
    End Property
    Public Property gobser1_3a
        Get
            Return obser1_3a
        End Get
        Set(value)
            obser1_3a = value
        End Set
    End Property
    Public Property gobser2_3a
        Get
            Return obser2_3a
        End Get
        Set(value)
            obser2_3a = value
        End Set
    End Property
    Public Property gobser3_3a
        Get
            Return obser3_3a
        End Get
        Set(value)
            obser3_3a = value
        End Set
    End Property
    Public Property gobser4_3a
        Get
            Return obser4_3a
        End Get
        Set(value)
            obser4_3a = value
        End Set
    End Property
    Public Property gobser5_3a
        Get
            Return obser5_3a
        End Get
        Set(value)
            obser5_3a = value
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
    Public Property gcanencio
        Get
            Return canencio
        End Get
        Set(value)
            canencio = value
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
    Public Property glote
        Get
            Return lote
        End Get
        Set(value)
            lote = value
        End Set
    End Property
    Public Property gunidad_medida
        Get
            Return unidad_medida
        End Get
        Set(value)
            unidad_medida = value
        End Set
    End Property
    Public Property gunidad_medida2
        Get
            Return unidad_medida2
        End Get
        Set(value)
            unidad_medida2 = value
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
    Public Property gitems
        Get
            Return items
        End Get
        Set(value)
            items = value
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
    Public Property gordens
        Get
            Return ordens
        End Get
        Set(value)
            ordens = value
        End Set
    End Property
    Public Property gparti
        Get
            Return parti
        End Get
        Set(value)
            parti = value
        End Set
    End Property
    Public Property grollo
        Get
            Return rollo
        End Get
        Set(value)
            rollo = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal sfactu As String,
    ByVal nfactu As String,
        ByVal items As String,
        ByVal linea As String,
        ByVal cantidad As Double, ByVal clasif As String,
        ByVal ordens As String, ByVal ccia As String,
        ByVal parti As String, ByVal lote As String, ByVal ordtejeduria As String,
        ByVal rollo As Double, ByVal unidad_medida As String, ByVal unidad_medida2 As String, ByVal canencio As Double, ByVal obser1_3a As String,
    ByVal obser2_3a As String,
        ByVal obser3_3a As String,
        ByVal obser4_3a As String,
        ByVal obser5_3a As String)
        gclasif = clasif
        gobser1_3a = obser1_3a
        gobser2_3a = obser2_3a
        gobser3_3a = obser3_3a
        gobser4_3a = obser4_3a
        gobser5_3a = obser5_3a
        gsfactu = sfactu
        gnfactu = nfactu
        gitems = items
        glinea = linea
        gcantidad = cantidad
        gordens = ordens
        gparti = parti
        grollo = rollo
        gunidad_medida2 = unidad_medida2
        gunidad_medida = unidad_medida
        glote = lote
        gccia = ccia
        gcanencio = canencio
        gordtejeduria = ordtejeduria
    End Sub
End Class
