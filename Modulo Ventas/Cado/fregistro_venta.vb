Imports System.Data.SqlClient
Public Class fregistro_venta
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function motrar_rsocial(ByVal dts As vregistroventa)
        Try
            conectare()
            cmd = New SqlCommand("cliente_ruc")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("nomb_10").ToString()
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
    Public Function motrar_rsocial2(ByVal dts As vregistroventa) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("cliente_ruc")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)

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
    Public Function motrar_rsocial3(ByVal dts As vregistroventa) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("cliente_ruc")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)

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
    Public Function insertar_cabecera_VENTA(ByVal dts As vregistroventa) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_VENTA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cperiodo_3", dts.gcperiodo_3)
            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom_3)
            cmd.Parameters.AddWithValue("@tidoc_3", dts.gtidoc_3)
            cmd.Parameters.AddWithValue("@sfactu_3", dts.gsfactu_3)
            cmd.Parameters.AddWithValue("@nfactu_3", dts.gnfactu_3)
            cmd.Parameters.AddWithValue("@fcom_3", dts.gfcom_3)
            cmd.Parameters.AddWithValue("@condp_3", dts.gcondp_3)
            cmd.Parameters.AddWithValue("@fich_3", dts.gfich_3)
            cmd.Parameters.AddWithValue("@nomb_3", dts.gnomb_3)
            cmd.Parameters.AddWithValue("@mone_3", dts.gmone_3)
            cmd.Parameters.AddWithValue("@tcam_3", dts.gtcam_3)
            cmd.Parameters.AddWithValue("@pvta1_3", dts.gpvta1_3)
            cmd.Parameters.AddWithValue("@vvta1_3", dts.gvvta1_3)
            cmd.Parameters.AddWithValue("@igv1_3", dts.gigv1_3)
            cmd.Parameters.AddWithValue("@tot1_3", dts.gtot1_3)
            cmd.Parameters.AddWithValue("@pvta2_3", dts.gpvta2_3)
            cmd.Parameters.AddWithValue("@vvta2_3", dts.gvvta2_3)
            cmd.Parameters.AddWithValue("@igv2_3", dts.gigv2_3)
            cmd.Parameters.AddWithValue("@tot2_3", dts.gtot2_3)
            cmd.Parameters.AddWithValue("@gloa_3", dts.ggloa_3)
            cmd.Parameters.AddWithValue("@aigv_3", dts.gaigv_3)
            cmd.Parameters.AddWithValue("@iigv_3", dts.giigv_3)
            cmd.Parameters.AddWithValue("@anal1_3", dts.ganal1_3)
            cmd.Parameters.AddWithValue("@porven_3", dts.gporven_3)
            cmd.Parameters.AddWithValue("@flag_3", dts.gflag_3)
            cmd.Parameters.AddWithValue("@ocompra_3", dts.gocompra_3)
            cmd.Parameters.AddWithValue("@vendedor_3", dts.gvendedor_3)
            cmd.Parameters.AddWithValue("@porcigv_3", dts.gporcigv_3)
            cmd.Parameters.AddWithValue("@tipo_venta", dts.gtipo_venta)
            cmd.Parameters.AddWithValue("@fecha_pago", dts.gfechacan)
            cmd.Parameters.AddWithValue("@aelanto", dts.gadeelanto)
            cmd.Parameters.AddWithValue("@fecha_cpago", dts.gfecha_cpago)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)
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
    Public Function insertar_detalle_VENTA(ByVal dts As vdetalleregistro) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_detalle_venta")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cperiodo_3a", dts.gcperiodo_3a)
            cmd.Parameters.AddWithValue("@ncom_3a", dts.gncom_3a)
            cmd.Parameters.AddWithValue("@item_3a", dts.gitem_3a)
            cmd.Parameters.AddWithValue("@linea_3a", dts.glinea_3a)
            cmd.Parameters.AddWithValue("@producto", dts.gproducto)
            cmd.Parameters.AddWithValue("@cant_3a", dts.gcant_3a)
            cmd.Parameters.AddWithValue("@unid_3a", dts.gunid_3a)
            cmd.Parameters.AddWithValue("@preun1_3a", dts.gpreun1_3a)
            cmd.Parameters.AddWithValue("@pvta1_3a", dts.gpvta1_3a)
            cmd.Parameters.AddWithValue("@vvta1_3a", dts.gvvta1_3a)
            cmd.Parameters.AddWithValue("@igv1_3a", dts.gigv1_3a)
            cmd.Parameters.AddWithValue("@tot1_3a", dts.gtot1_3a)
            cmd.Parameters.AddWithValue("@preun2_3a", dts.gpreun2_3a)
            cmd.Parameters.AddWithValue("@pvta2_3a", dts.gpvta2_3a)
            cmd.Parameters.AddWithValue("@vvta2_3a", dts.gvvta2_3a)
            cmd.Parameters.AddWithValue("@igv2_3a", dts.gigv2_3a)
            cmd.Parameters.AddWithValue("@tot2_3a", dts.gtot2_3a)
            cmd.Parameters.AddWithValue("@obser1_3a", dts.gobser1_3a)
            cmd.Parameters.AddWithValue("@porven_3a", dts.gporven_3a)
            cmd.Parameters.AddWithValue("@ordenp_3a", dts.gordenp_3a)
            cmd.Parameters.AddWithValue("@grm_3a", dts.ggrm_3a)
            cmd.Parameters.AddWithValue("@flag_3a", dts.gflag_3a)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)
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
    Public Function mostrar_venta(ByVal dts As vregistroventa) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_ventasn")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@estado", dts.ganal1_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function mostrar_detalle_venta(ByVal dts As vregistroventa) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_det_venta")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@id", dts.gncom_3)
            cmd.Parameters.AddWithValue("@periodo", dts.gcperiodo_3)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function actualizar_venta(ByVal dts As vfactura) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_factura_ventas")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function actualizar_venta_MONTO(ByVal dts As vfactura) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("MONTO_PARCIAL_ventas")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@MONTO", dts.gMONTO)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function existeregistrovn(ByVal dts As vdetalleregistro)
        Try
            conectare()
            cmd = New SqlCommand("exite_ventan")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom_3a)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
            cmd.Parameters.AddWithValue("@periodo", dts.gcperiodo_3a)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("can").ToString()
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
    Public Function buscar_ventas_negra_cabecera(ByVal dts As vdetalleregistro) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_ventasncab")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ncom", dts.gncom_3a)
            cmd.Parameters.AddWithValue("@periodo", dts.gcperiodo_3a)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function buscar_ventas_negra_DETALLE(ByVal dts As vdetalleregistro) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_ventasndeta")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ncom", dts.gncom_3a)
            cmd.Parameters.AddWithValue("@periodo", dts.gcperiodo_3a)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function eliminar_regventanegr(ByVal dts As vdetalleregistro) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_ventan")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom_3a)
            cmd.Parameters.AddWithValue("@periodo", dts.gcperiodo_3a)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
