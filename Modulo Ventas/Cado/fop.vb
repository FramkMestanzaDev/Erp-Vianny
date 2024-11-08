Imports System.Data.SqlClient
Public Class fop
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_op(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_op")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom_3)
            cmd.Parameters.AddWithValue("@fich_3", dts.gfich_3)
            cmd.Parameters.AddWithValue("@fcom_3", dts.gfcom_3)
            cmd.Parameters.AddWithValue("@mone_3", dts.gmone_3)
            cmd.Parameters.AddWithValue("@linea_3", dts.glinea_3)
            cmd.Parameters.AddWithValue("@FCome1_3", dts.gFCome1_3)
            cmd.Parameters.AddWithValue("@frequerida_3", dts.gfrequerida_3)
            cmd.Parameters.AddWithValue("@tcam_3", dts.gtcam_3)
            cmd.Parameters.AddWithValue("@program_3", dts.gprogram_3)
            cmd.Parameters.AddWithValue("@flag_3", dts.gflag_3)
            cmd.Parameters.AddWithValue("@descri_3", dts.gdescri_3)
            cmd.Parameters.AddWithValue("@alterno_3", dts.galterno_3)
            cmd.Parameters.AddWithValue("@estilo_3", dts.gestilo_3)
            cmd.Parameters.AddWithValue("@pfob_3", dts.gpfob_3)
            cmd.Parameters.AddWithValue("@cantp_3", dts.gcantp_3)
            cmd.Parameters.AddWithValue("@cants_3", dts.gcants_3)
            cmd.Parameters.AddWithValue("@merchan_3", dts.gmerchan_3)
            cmd.Parameters.AddWithValue("@broker_3", dts.gbroker_3)
            cmd.Parameters.AddWithValue("@tela1_3", dts.gtela1_3)
            cmd.Parameters.AddWithValue("@telaprinc_3", dts.gtelaprinc_3)
            cmd.Parameters.AddWithValue("@fpago_3", dts.gfpago_3)
            cmd.Parameters.AddWithValue("@ffinprod_3", dts.gffinprod_3)
            cmd.Parameters.AddWithValue("@observ_3", dts.gobserv_3)
            cmd.Parameters.AddWithValue("@cod_color", dts.gcod_color)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@familia", dts.gfamilia)
            cmd.Parameters.AddWithValue("@estiloemp_3", dts.gestiloemp_3)
            cmd.Parameters.AddWithValue("@ID_REQUERIMIENTO", dts.gID_REQUERIMIENTO)
            'cmd.Parameters.AddWithValue("@ttela", dts.gttela)
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

    Public Function insertar_cab_consumo(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("FichaConsumoTextilCabInsertUpdated")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
            cmd.Parameters.AddWithValue("@FEMISION", dts.gfcom_3)
            cmd.Parameters.AddWithValue("@nTIPO", dts.gtipo)

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
    Public Function insertar_detelle_opp(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_detelle_opp")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@op", dts.gncom_3)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantp_3)

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
    Public Function insertar_cdet_consumo(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("FichaConsumoTextilDetInsertUpdated")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
            cmd.Parameters.AddWithValue("@TELA", dts.glinea_3)


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
    Public Function actualizar_OP(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ANULAR_OP")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function actualizar_op_obser(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_op_obser")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@obser", dts.goobservacion2)
            cmd.Parameters.AddWithValue("@op", dts.gncom_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function cerrar_op(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("cerrar_op")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)

            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function standby_op(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("standby_op")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)

            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function ELIMINAR_OP(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ELIMINAR_OP")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function ELIMINAR_OP2(ByVal dts As vop) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ELIMINAR_OP_2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
            cmd.Parameters.AddWithValue("@CODIGO", dts.glinea_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function PROGRAMA_TEJEDURIA() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_PROGRAMA_TEJEDURIA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

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
    Public Function BUSCAR_PROGRAMA_standby() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_PROGRAMA_standby")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

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
    Public Function PROGRAMA_TEJEDURIA_CERRADO() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_PROGRAMA_TEJEDURIA_CERRADO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

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
    Public Function mostarr_op(ByVal dts As vop) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_OP_COMERCIAL")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
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
    Public Function buscar_op(ByVal dts As vop) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("op")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@PO", dts.gncom_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
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
    Public Function PADRON_TALLA(ByVal dts As vop) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("talla_op")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
            cmd.Parameters.AddWithValue("@CCIA", dts.gcia)
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
    Public Function PADRON_TALLA1(ByVal dts As vop) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("talla_op1")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@OP", dts.gncom_3)
            cmd.Parameters.AddWithValue("@CCIA", dts.gcia)
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
    Public Function validar_op(ByVal dts As vop)
        Try
            conectare()
            cmd = New SqlCommand("validar_op")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gncom_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("cant").ToString()
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function buscar_observacion(ByVal dts As vop)
        Try
            conectare()
            cmd = New SqlCommand("buscar_observacion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gncom_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("estobs_3").ToString()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function liquidar_produccion(ByVal dts As vop)
        Try
            conectare()
            cmd = New SqlCommand("liquidar_produccion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gncom_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
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
End Class
