Imports System.Data.SqlClient
Public Class ftejeduria
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function eliminar_defectos(ByVal dts As vtejeduria) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.mvdefectos_rollo_delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@tdr_ccia", dts.gtdr_ccia)
            cmd.Parameters.AddWithValue("@tsdr_rollo", dts.gtsdr_rollo)
            cmd.Parameters.AddWithValue("@tdr_defecto", dts.gtdr_defecto)



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
    Public Function actualizar_defectos(ByVal dts As vtejeduria) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.mvdefectos_rollo_update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@tdr_ccia", dts.gtdr_ccia)
            cmd.Parameters.AddWithValue("@tsdr_rollo", dts.gtsdr_rollo)
            cmd.Parameters.AddWithValue("@tdr_defecto", dts.gtdr_defecto)
            cmd.Parameters.AddWithValue("@tdr_obs", dts.gtdr_obs)
            cmd.Parameters.AddWithValue("@tdr_cantidad", dts.gtdr_cantidad)
            cmd.Parameters.AddWithValue("@tdr_mtrafectados", dts.gtdr_mtrafectados)
            cmd.Parameters.AddWithValue("@tdr_kgafectados", dts.gtdr_kgafectados)


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
    Public Function insertar_defectos(ByVal dts As vtejeduria) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_defectos_rollos")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@tdr_codigo", dts.gtdr_codigo)
            cmd.Parameters.AddWithValue("@tdr_descripcion", dts.gtdr_descripcion)
            cmd.Parameters.AddWithValue("@tdr_usuario", dts.gtdr_usuario)
            cmd.Parameters.AddWithValue("@tdr_fupdate", dts.gtdr_fupdate)
            cmd.Parameters.AddWithValue("@tdr_pc", dts.gtdr_pc)
            cmd.Parameters.AddWithValue("@tdr_tipodefecto", dts.gtdr_tipodefecto)


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
    Public Function reporte_pesado_rollo(ByVal dts As vtejeduria) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.marg0001_consulta")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@rolloini", dts.grolloini)
            cmd.Parameters.AddWithValue("@rollofin", dts.grollofin)
            cmd.Parameters.AddWithValue("@pedidoot", DBNull.Value)
            cmd.Parameters.AddWithValue("@correlativo", DBNull.Value)
            cmd.Parameters.AddWithValue("@ordetejido", DBNull.Value)
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

    Public Function reporte_tejeduria(ByVal dts As vtejeduria) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("SPU_ProduccionTejeduria")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@FECHAINI", dts.gfechaini)
            cmd.Parameters.AddWithValue("@FECHAFIN", dts.gfechafin)
            If dts.gtejedor = "" Then
                cmd.Parameters.AddWithValue("@TEJEDOR", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@TEJEDOR", dts.gtejedor)
            End If
            If dts.gmaquina = "" Then
                cmd.Parameters.AddWithValue("@MAQUINA", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@MAQUINA", dts.gmaquina)
            End If

            cmd.Parameters.AddWithValue("@TIPO_TELA", DBNull.Value)
            cmd.Parameters.AddWithValue("@TITULO", DBNull.Value)
            cmd.Parameters.AddWithValue("@TURNO", DBNull.Value)
            cmd.Parameters.AddWithValue("@NQUIEBRE", 1)
            cmd.Parameters.AddWithValue("@NMODALIDAD", 1)
            cmd.Parameters.AddWithValue("@NAGRUPADO", 1)
            cmd.Parameters.AddWithValue("@CLIENTE", DBNull.Value)
            cmd.Parameters.AddWithValue("@ORDTEJ", DBNull.Value)
            If dts.gpedidoot = "" Then
                cmd.Parameters.AddWithValue("@PEDIDOINI", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@PEDIDOINI", dts.gpedidoot)
            End If
            If dts.gpedidoot = "" Then
                cmd.Parameters.AddWithValue("@PEDIDOFIN", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@PEDIDOFIN", dts.gpedidoot)
            End If

            cmd.Parameters.AddWithValue("@CODTEJEDURIA", DBNull.Value)
            cmd.Parameters.AddWithValue("@LOTE", DBNull.Value)
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
    Public Function reporte_tejeduria2(ByVal dts As vtejeduria) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("prduccion_tejeduria")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@FECHAINI", dts.gfechaini)
            cmd.Parameters.AddWithValue("@FECHAFIN", dts.gfechafin)

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
    Public Function reporte_calidad(ByVal dts As vtejeduria) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_CALIDAD_ROLLOS")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            If dts.gtejedor = "" Then
                cmd.Parameters.AddWithValue("@PARTIDA", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@PARTIDA", dts.gpartida)
            End If
            If dts.gmaquina = "" Then
                cmd.Parameters.AddWithValue("@PEDIDO", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@PEDIDO", dts.gpo)
            End If
            cmd.Parameters.AddWithValue("@FECHAINI", dts.gfechaini)
            cmd.Parameters.AddWithValue("@FECHAFIN", dts.gfechafin)

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
    Public Function reporte_defectos_rollo(ByVal dts As vtejeduria) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("spu_cargadefectosxrolloj")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@pedidoot", DBNull.Value)
            cmd.Parameters.AddWithValue("@correlativo", DBNull.Value)
            cmd.Parameters.AddWithValue("@ordetejido", DBNull.Value)
            cmd.Parameters.AddWithValue("@rolloini", dts.grolloini)
            cmd.Parameters.AddWithValue("@rollofin", dts.grollofin)
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
    Public Function buscar_ventas_m(ByVal dts As vtejeduria) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_ventas_m")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fechaini", dts.gfechaini)
            cmd.Parameters.AddWithValue("@fechafin", dts.gfechafin)

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
    Public Function insertar_pesado_rollos(ByVal dts As vpesadorollo) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_pesado_rollo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia_3r", dts.gccia_3r)
            cmd.Parameters.AddWithValue("@prvtej_3r", dts.gprvtej_3r)
            cmd.Parameters.AddWithValue("@rollo_3r", dts.grollo_3r)
            cmd.Parameters.AddWithValue("@rollom_3r", dts.grollom_3r)
            cmd.Parameters.AddWithValue("@ncom_3r", dts.gncom_3r)
            cmd.Parameters.AddWithValue("@pedido_3r", dts.gpedido_3r)
            cmd.Parameters.AddWithValue("@galga_3r", dts.ggalga_3r)
            cmd.Parameters.AddWithValue("@lote_3r", dts.glote_3r)
            cmd.Parameters.AddWithValue("@linea_3r", dts.glinea_3r)
            cmd.Parameters.AddWithValue("@maquina_3r", dts.gmaquina_3r)
            cmd.Parameters.AddWithValue("@cantk_3r", dts.gcantk_3r)
            cmd.Parameters.AddWithValue("@prvhilo_3r", dts.gprvhilo_3r)
            cmd.Parameters.AddWithValue("@fcom_3r", dts.gfcom_3r)
            cmd.Parameters.AddWithValue("@fproduc_3r", dts.gfproduc_3r)
            cmd.Parameters.AddWithValue("@partida_3r", dts.gpartida_3r)
            cmd.Parameters.AddWithValue("@turno_3r", dts.gturno_3r)
            cmd.Parameters.AddWithValue("@grm_3r", dts.ggrm_3r)
            cmd.Parameters.AddWithValue("@situa_3r", dts.gsitua_3r)
            cmd.Parameters.AddWithValue("@flag_3r", dts.gflag_3r)
            cmd.Parameters.AddWithValue("@pedmov_3r", dts.gpedmov_3r)
            cmd.Parameters.AddWithValue("@cantkmv_3r", dts.gcantkmv_3r)
            cmd.Parameters.AddWithValue("@ancho_3r", dts.gancho_3r)
            cmd.Parameters.AddWithValue("@diam_3r", dts.gdiam_3r)
            cmd.Parameters.AddWithValue("@usokar_3r", dts.gusokar_3r)
            cmd.Parameters.AddWithValue("@ordens_3r", dts.gordens_3r)
            cmd.Parameters.AddWithValue("@ccolor_3r", dts.gccolor_3r)
            cmd.Parameters.AddWithValue("@calidad_3r", dts.gcalidad_3r)
            cmd.Parameters.AddWithValue("@ubica_3r", dts.gubica_3r)
            cmd.Parameters.AddWithValue("@ordtej_3r", dts.gordtej_3r)
            cmd.Parameters.AddWithValue("@fecgen_3r", dts.gfecgen_3r)
            cmd.Parameters.AddWithValue("@tejedor_3r", dts.gtejedor_3r)
            cmd.Parameters.AddWithValue("@tiprol_3r", dts.gtiprol_3r)
            cmd.Parameters.AddWithValue("@solmue_3r", dts.gsolmue_3r)
            cmd.Parameters.AddWithValue("@prvtin_3r", dts.gprvtin_3r)
            cmd.Parameters.AddWithValue("@almacen_3r", dts.galmacen_3r)
            cmd.Parameters.AddWithValue("@periodo_3r", dts.gperiodo_3r)
            cmd.Parameters.AddWithValue("@voucher_3r", dts.gvoucher_3r)
            cmd.Parameters.AddWithValue("@zarmado_3r", dts.gzarmado_3r)
            cmd.Parameters.AddWithValue("@letradiv_3r", dts.gletradiv_3r)
            cmd.Parameters.AddWithValue("@mtsafec_3a", dts.gmtsafec_3a)
            cmd.Parameters.AddWithValue("@kgsafec_3a", dts.gkgsafec_3a)
            cmd.Parameters.AddWithValue("@calrol_3r", dts.gcalrol_3r)
            cmd.Parameters.AddWithValue("@rendimiento_3r", dts.grendimiento_3r)
            cmd.Parameters.AddWithValue("@obscalidad_3r", dts.gobscalidad_3r)
            cmd.Parameters.AddWithValue("@retrol_3r", dts.gretrol_3r)
            cmd.Parameters.AddWithValue("@escalrol_3r", dts.gescalrol_3r)
            cmd.Parameters.AddWithValue("@densidad_3r", dts.gdensidad_3r)
            cmd.Parameters.AddWithValue("@audicrudo_3r", dts.gaudicrudo_3r)
            cmd.Parameters.AddWithValue("@frevcrudo_3r", dts.gfrevcrudo_3r)
            cmd.Parameters.AddWithValue("@anchor_3r", dts.ganchor_3r)
            cmd.Parameters.AddWithValue("@densidadr_3r", dts.gdensidadr_3r)
            cmd.Parameters.AddWithValue("@rendimientor_3r", dts.grendimientor_3r)
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
    Public Function MOSTRAR_TEJEDOR(ByVal DTS As vpesadorollo)
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_TEJEDOR")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@TEJEDOR", DTS.gtejedor)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("nomb_10").ToString()
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
End Class
