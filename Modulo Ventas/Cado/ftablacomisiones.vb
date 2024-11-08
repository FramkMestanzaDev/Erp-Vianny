Imports System.Data.SqlClient
Public Class ftablacomisiones
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_tabal_comisiones(ByVal dts As vtablacomisiones) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_tabal_comisiones")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@rubro", dts.grubro)
            cmd.Parameters.AddWithValue("@fpagocontado", dts.gfpagocontado)
            cmd.Parameters.AddWithValue("@preciocontado", dts.gpreciocontado)
            cmd.Parameters.AddWithValue("@operacioncontado", dts.goperacioncontado)
            cmd.Parameters.AddWithValue("@comisioncontado", dts.gcomisioncontado)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@id", dts.gid)
            cmd.Parameters.AddWithValue("@rubro_codigo", dts.grubro_codigo)
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
    Public Function buscar_tabal_comisione(ByVal dts As vtablacomisiones) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_tabal_comisiones")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            'cmd.Parameters.AddWithValue("@mes", dts.gmes)

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
