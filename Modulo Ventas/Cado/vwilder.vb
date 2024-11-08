Imports System.Data.SqlClient
Public Class vwilder
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar(ByVal dts As fwilder) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_registro_almacen")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@op", dts.gop)
            cmd.Parameters.AddWithValue("@cotizacion", dts.gcotizacion)
            cmd.Parameters.AddWithValue("@nia", dts.gnia)
            cmd.Parameters.AddWithValue("@dir_nia", dts.gdir_nia)
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
            cmd.Parameters.AddWithValue("@dir_fac", dts.gdir_fac)
            cmd.Parameters.AddWithValue("@guia", dts.gguia)
            cmd.Parameters.AddWithValue("@dir_guia", dts.gdir_guia)
            cmd.Parameters.AddWithValue("@ocompras", dts.gocompras)
            cmd.Parameters.AddWithValue("@dir_ocompra", dts.gdir_ocompra)
            cmd.Parameters.AddWithValue("@lacrado", dts.glacrado)
            cmd.Parameters.AddWithValue("@lacrado_contabilidad", dts.glacrado_contab)
            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox("LA FACTURA YA ESTA REGISTRADA")
            Return False
        Finally
            desconectar()
        End Try
    End Function
    Public Function eliminar(ByVal dts As fwilder) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_registro_almacen")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
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
    Public Function buscar() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_registro_almacen")
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
    Public Function buscar_contabilidad() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCARr_contabilidad")
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
End Class
