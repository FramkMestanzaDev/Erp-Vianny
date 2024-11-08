Public Class vprogramacion
    Dim op As String
    Dim codigo As String
    Dim fecha As Date
    Dim color As String
    Dim codigo1 As String
    Dim codigo2 As String
    Dim codigo3 As String
    Dim valor2 As Double
    Dim valor1 As Double
    Dim valor3 As Double
    Dim valor As Double
    Dim valor4 As Double
    Dim ccia As String

    Public Property gop
        Get
            Return op
        End Get
        Set(value)
            op = value
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
    Public Property gfecha
        Get
            Return fecha
        End Get
        Set(value)
            fecha = value
        End Set
    End Property
    Public Property gcolor
        Get
            Return color
        End Get
        Set(value)
            color = value
        End Set
    End Property
    Public Property gcodigo1
        Get
            Return codigo1
        End Get
        Set(value)
            codigo1 = value
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
    Public Property gcodigo3
        Get
            Return codigo3
        End Get
        Set(value)
            codigo3 = value
        End Set
    End Property
    Public Property gvalor2
        Get
            Return valor2
        End Get
        Set(value)
            valor2 = value
        End Set
    End Property
    Public Property gvalor1
        Get
            Return valor1
        End Get
        Set(value)
            valor1 = value
        End Set
    End Property
    Public Property gvalor3
        Get
            Return valor3
        End Get
        Set(value)
            valor3 = value
        End Set
    End Property
    Public Property gvalor
        Get
            Return valor
        End Get
        Set(value)
            valor = value
        End Set
    End Property
    Public Property gvalor4
        Get
            Return valor4
        End Get
        Set(value)
            valor4 = value
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
    Sub New()

    End Sub
    Sub New(ByVal op As String, ByVal codigo As String, ByVal fecha As Date, ByVal color As String, ByVal codigo1 As String, ByVal codigo2 As String, ByVal codigo3 As String, ByVal valor2 As Double,
            ByVal valor1 As Double, ByVal valor3 As Double, ByVal valor As Double, ByVal valor4 As Double, ByVal ccia As String)
        gop = op
        gcodigo = codigo
        gfecha = fecha
        gcolor = color
        gcodigo1 = codigo1
        gcodigo2 = codigo2
        gcodigo3 = codigo3
        gvalor2 = valor2
        gvalor1 = valor1
        gvalor3 = valor3
        gvalor = valor
        gvalor4 = valor4
        gccia = ccia

    End Sub
End Class
