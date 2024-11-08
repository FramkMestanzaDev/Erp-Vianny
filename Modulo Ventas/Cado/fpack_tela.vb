Imports System.Data.SqlClient
Public Class fpack_tela
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function correlativo_pack_tela()
        Try
            conectare()
            cmd = New SqlCommand("correlativo_packung_tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("VAL").ToString()
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
    Public Function insertar_packin_tela(ByVal dts As vpackin_tela_cab) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_pcking_tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@id_packing", dts.gid_packing)
            cmd.Parameters.AddWithValue("@cod_pedido", dts.gcod_pedido)
            cmd.Parameters.AddWithValue("@nota_pedido", dts.gnota_pedido)
            cmd.Parameters.AddWithValue("@codigo_pro", dts.gcodigo_pro)
            cmd.Parameters.AddWithValue("@descrip_pro", dts.gdescrip_pro)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@f_despacho", dts.gf_despacho)
            cmd.Parameters.AddWithValue("@f_actual", dts.gf_actual)
            cmd.Parameters.AddWithValue("@cliente", dts.gcliente)
            cmd.Parameters.AddWithValue("@Vendedor", dts.gVendedor)

            cmd.Parameters.AddWithValue("@kg_seleccionados", dts.gkg_seleccionados)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
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
    Public Function insertar_packin_tela_DETALLE(ByVal dts As vpackin_tela_det) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_DETALLE_PACKING")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@rollo", dts.grollo)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@ubicacion", dts.gubicacion)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@id_packing", dts.gid_packing)
            cmd.Parameters.AddWithValue("@codigo_pro", dts.gcodigo_pro)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
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
