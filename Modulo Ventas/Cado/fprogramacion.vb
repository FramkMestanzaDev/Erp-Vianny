Imports System.Data.SqlClient
Public Class fprogramacion
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscar_tejido(ByVal dts As vprogramacion) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.CaeSoft_GetAllFichaTextilTelaDetalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CODIGO_TELA", dts.gcodigo)
            If cmd.ExecuteNonQuery Then
                Dim dt As New DataTable
                Dim da As New SqlDataAdapter(cmd)
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function buscar_FICHA(ByVal dts As vprogramacion) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.GetallFichaConsumoTextilDet")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@OP", dts.gop)
            If cmd.ExecuteNonQuery Then
                Dim dt As New DataTable
                Dim da As New SqlDataAdapter(cmd)
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function buscar_ot_partida(ByVal dts As vprogramacion) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.CaeSoft_RequerimientoProgramadoPedidoCrudoDetallado")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gccia)
            cmd.Parameters.AddWithValue("@Pedido", dts.gop)
            If cmd.ExecuteNonQuery Then
                Dim dt As New DataTable
                Dim da As New SqlDataAdapter(cmd)
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function insertar_Qarg0300(ByVal dts As vprogramacion) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_Qarg0300")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@op", dts.gop)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)


            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
    Public Function insertar_Qarp0300(ByVal dts As vprogramacion) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_Qarp0300")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@op", dts.gop)
            'cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            'cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@codigo1", dts.gcodigo1)
            'cmd.Parameters.AddWithValue("@codigo2", dts.gcodigo2)
            'cmd.Parameters.AddWithValue("@codigo3", dts.gcodigo3)
            'cmd.Parameters.AddWithValue("@valor2", dts.gvalor2)
            cmd.Parameters.AddWithValue("@valor1", dts.gvalor1)
            'cmd.Parameters.AddWithValue("@valor3", dts.gvalor3)
            'cmd.Parameters.AddWithValue("@valor", dts.gvalor)
            'cmd.Parameters.AddWithValue("@valor4", dts.gvalor4)

            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
    Public Function insertar_Qarp0300_det(ByVal dts As vprogramacion) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_Qarp0300d")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@op", dts.gop)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)

            'cmd.Parameters.AddWithValue("@codigo2", dts.gcodigo2)
            'cmd.Parameters.AddWithValue("@codigo3", dts.gcodigo3)
            'cmd.Parameters.AddWithValue("@valor2", dts.gvalor2)

            'cmd.Parameters.AddWithValue("@valor3", dts.gvalor3)
            cmd.Parameters.AddWithValue("@valor", dts.gvalor)
            cmd.Parameters.AddWithValue("@valor4", dts.gvalor4)

            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
    'Public Function insertar_matg040f(ByVal dts As vmatg040f) As Boolean
    '    Try
    '        conectare()
    '        cmd = New SqlCommand("insertar_matg040f")
    '        cmd.CommandType = CommandType.StoredProcedure
    '        cmd.Connection = cnn


    '        cmd.Parameters.AddWithValue("@ccia_4", dts.gccia_4)
    '        cmd.Parameters.AddWithValue("@ncom4", dts.gncom4)
    '        cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
    '        cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
    '        cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
    '        cmd.Parameters.AddWithValue("@partida_4", dts.gpartida_4)
    '        cmd.Parameters.AddWithValue("@color_4", dts.gcolor_4)
    '        cmd.Parameters.AddWithValue("@prvtin_4", dts.gprvtin_4)
    '        cmd.Parameters.AddWithValue("@prvtej_4", dts.gprvtej_4)
    '        cmd.Parameters.AddWithValue("@acabado_4", dts.gacabado_4)
    '        cmd.Parameters.AddWithValue("@presen_4", dts.gpresen_4)
    '        cmd.Parameters.AddWithValue("@psxrolst_4", dts.gpsxrolst_4)
    '        cmd.Parameters.AddWithValue("@nrevrol_4", dts.gnrevrol_4)
    '        cmd.Parameters.AddWithValue("@rpmrol_4", dts.grpmrol_4)

    '        If cmd.ExecuteNonQuery Then
    '            Return True
    '        Else
    '            Return False
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '        Return False
    '    Finally
    '        desconectar()
    '    End Try
    'End Function
    Public Function insertar_matg040f(ByVal dts As vmatg040f) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_matg040f")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia_4", dts.gccia_4)
            cmd.Parameters.AddWithValue("@ncom4", dts.gncom4)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
            cmd.Parameters.AddWithValue("@partida_4", dts.gpartida_4)
            cmd.Parameters.AddWithValue("@color_4", dts.gcolor_4)
            cmd.Parameters.AddWithValue("@prvtin_4", dts.gprvtin_4)
            cmd.Parameters.AddWithValue("@prvtej_4", dts.gprvtej_4)
            cmd.Parameters.AddWithValue("@acabado_4", dts.gacabado_4)
            cmd.Parameters.AddWithValue("@presen_4", dts.gpresen_4)
            cmd.Parameters.AddWithValue("@psxrolst_4", dts.gpsxrolst_4)
            cmd.Parameters.AddWithValue("@nrevrol_4", dts.gnrevrol_4)
            cmd.Parameters.AddWithValue("@rpmrol_4", dts.grpmrol_4)

            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
    Public Function insertar_matp040f(ByVal dts As vmatp040f) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_matp040f")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia_4a", dts.gccia_4a)
            cmd.Parameters.AddWithValue("@ncom_4a", dts.gncom_4a)
            cmd.Parameters.AddWithValue("@TELA_4A", dts.gTELA_4A)
            cmd.Parameters.AddWithValue("@DIAMET_4A", dts.gDIAMET_4A)
            cmd.Parameters.AddWithValue("@GALGA_4A", dts.gGALGA_4A)
            cmd.Parameters.AddWithValue("@CANTREQ_4A", dts.gCANTREQ_4A)
            cmd.Parameters.AddWithValue("@MAQUINA_4A", dts.gMAQUINA_4A)
            cmd.Parameters.AddWithValue("@ANCHO_4A", dts.gANCHO_4A)
            cmd.Parameters.AddWithValue("@LM1_41A", dts.gLM1_41A)
            cmd.Parameters.AddWithValue("@LM2_41A", dts.gLM2_41A)
            cmd.Parameters.AddWithValue("@LM3_41A", dts.gLM3_41A)
            cmd.Parameters.AddWithValue("@LM4_41A", dts.gLM4_41A)
            cmd.Parameters.AddWithValue("@LM5_41A", dts.gLM5_41A)


            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
    Public Function insertar_matp041f(ByVal dts As vmatp041f) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_matp041f")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia_41A", dts.gccia_41A)
            cmd.Parameters.AddWithValue("@ncom_41a", dts.gncom_41a)
            cmd.Parameters.AddWithValue("@itm_41a", dts.gitm_41a)
            cmd.Parameters.AddWithValue("@hilo_41a", dts.ghilo_41a)
            cmd.Parameters.AddWithValue("@cant_41a", dts.gcant_41a)
            cmd.Parameters.AddWithValue("@lote_41a", dts.glote_41a)
            cmd.Parameters.AddWithValue("@prvhil_41a", dts.gprvhil_41a)


            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
    Public Function insertar_qan0300(ByVal dts As vqan0300) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_qan0300")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia_3n", dts.gccia_3n)
            cmd.Parameters.AddWithValue("@regis_3n", dts.gregis_3n)
            cmd.Parameters.AddWithValue("@ncom_3n", dts.gncom_3n)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@origen_3n", dts.gorigen_3n)
            cmd.Parameters.AddWithValue("@Fich_3n", dts.gFich_3n)
            cmd.Parameters.AddWithValue("@color_3n", dts.gcolor_3n)
            cmd.Parameters.AddWithValue("@fase_3n", dts.gfase_3n)
            cmd.Parameters.AddWithValue("@estado_3n", dts.gestado_3n)
            cmd.Parameters.AddWithValue("@Glosa1_3n", dts.gGlosa1_3n)
            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
    Public Function insertar_qanp300(ByVal dts As vqanp300) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_qanp300")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia_3p", dts.gccia_3p)
            cmd.Parameters.AddWithValue("@regis_3p", dts.gregis_3p)
            cmd.Parameters.AddWithValue("@Op_3p", dts.gOp_3p)
            cmd.Parameters.AddWithValue("@linea_3p", dts.glinea_3p)
            cmd.Parameters.AddWithValue("@galga_3p", dts.ggalga_3p)
            cmd.Parameters.AddWithValue("@Ancho_3p", dts.gAncho_3p)
            cmd.Parameters.AddWithValue("@kgreq_3p", dts.gkgreq_3p)
            cmd.Parameters.AddWithValue("@observ_3p", dts.gobserv_3p)

            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            desconectar()
        End Try
    End Function
End Class
