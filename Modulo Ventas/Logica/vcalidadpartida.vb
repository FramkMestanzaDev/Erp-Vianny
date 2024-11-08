Public Class vcalidadpartida
    Dim partida As String
    Dim codigo As String
    Dim color_des As String
    Dim color_cod As String
    Dim descripcion As String
    Dim rollos As String
    Dim anchor As String
    Dim densidadr As String
    Dim lavadoa As String
    Dim lavadol As String
    Dim lavador As String
    Dim observacion As String
    Dim revisado As String
    Dim adjuntado As String
    Dim reproceso As String
    Dim estado As String
    Dim elongacion As String
    Dim fecha As Date
    Dim merma As Double
    Dim maquina As String
    Dim lote As String
    Dim CENTRO_ORILLO As String
    Dim DEGRADE As String
    Dim BARRADURA As String
    Public Property gCENTRO_ORILLO
        Get
            Return CENTRO_ORILLO
        End Get
        Set(value)
            CENTRO_ORILLO = value
        End Set
    End Property
    Public Property gDEGRADE
        Get
            Return DEGRADE
        End Get
        Set(value)
            DEGRADE = value
        End Set
    End Property
    Public Property gBARRADURA
        Get
            Return BARRADURA
        End Get
        Set(value)
            BARRADURA = value
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
    Public Property gmaquina
        Get
            Return maquina
        End Get
        Set(value)
            maquina = value
        End Set
    End Property
    Public Property gmerma
        Get
            Return merma
        End Get
        Set(value)
            merma = value
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
    Public Property gelongacion
        Get
            Return elongacion
        End Get
        Set(value)
            elongacion = value
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
    Public Property greproceso
        Get
            Return reproceso
        End Get
        Set(value)
            reproceso = value
        End Set
    End Property
    Public Property gadjuntado
        Get
            Return adjuntado
        End Get
        Set(value)
            adjuntado = value
        End Set
    End Property
    Public Property grevisado
        Get
            Return revisado
        End Get
        Set(value)
            revisado = value
        End Set
    End Property
    Public Property gobservacion
        Get
            Return observacion
        End Get
        Set(value)
            observacion = value
        End Set
    End Property
    Public Property glavador
        Get
            Return lavador
        End Get
        Set(value)
            lavador = value
        End Set
    End Property
    Public Property glavadol
        Get
            Return lavadol
        End Get
        Set(value)
            lavadol = value
        End Set
    End Property
    Public Property glavadoa
        Get
            Return lavadoa
        End Get
        Set(value)
            lavadoa = value
        End Set
    End Property
    Public Property gdensidadr
        Get
            Return densidadr
        End Get
        Set(value)
            densidadr = value
        End Set
    End Property
    Public Property ganchor
        Get
            Return anchor
        End Get
        Set(value)
            anchor = value
        End Set
    End Property
    Public Property grollos
        Get
            Return rollos
        End Get
        Set(value)
            rollos = value
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
    Public Property gcolor_cod
        Get
            Return color_cod
        End Get
        Set(value)
            color_cod = value
        End Set
    End Property
    Public Property gcolor_des
        Get
            Return color_des
        End Get
        Set(value)
            color_des = value
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
    Public Property gcodigo
        Get
            Return codigo
        End Get
        Set(value)
            codigo = value
        End Set
    End Property
    Sub New()

    End Sub
    Sub New(ByVal partida As String,
             ByVal codigo As String,
              ByVal color_des As String,
              ByVal color_cod As String,
              ByVal descripcion As String,
              ByVal rollos As String,
              ByVal anchor As String,
              ByVal densidadr As String,
              ByVal lavadoa As String,
              ByVal lavadol As String,
              ByVal lavador As String,
              ByVal observacion As String,
             ByVal revisado As String,
             ByVal adjuntado As String,
             ByVal reproceso As String,
             ByVal estado As String,
            ByVal elongacion As String,
            ByVal fecha As Date,
            ByVal merma As Double,
            ByVal maquina As String,
            ByVal lote As String,
             ByVal CENTRO_ORILLO As String,
    ByVal DEGRADE As String,
        ByVal BARRADURA As String
   )
        gCENTRO_ORILLO = CENTRO_ORILLO
        glote = lote
        gCENTRO_ORILLO = CENTRO_ORILLO
        gmaquina = maquina
        gmerma = merma
        gpartida = partida
        gcodigo = codigo
        gcolor_des = color_des
        gcolor_cod = color_cod
        gdescripcion = descripcion
        grollos = rollos
        ganchor = anchor
        gdensidadr = densidadr
        glavadoa = lavadoa
        glavadol = lavadol
        glavador = lavador
        gobservacion = observacion
        grevisado = revisado
        gadjuntado = adjuntado
        greproceso = reproceso
        gestado = estado
        gelongacion = elongacion
        gfecha = fecha
        glote = lote
    End Sub
End Class
