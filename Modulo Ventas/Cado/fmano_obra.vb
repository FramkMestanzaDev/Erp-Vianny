Imports System.Data.SqlClient
Public Class fmano_obra
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_mano_obra(ByVal dts As vManoObra) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_ManoObra")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@merma", dts.gmerma)
            cmd.Parameters.AddWithValue("@consumo", dts.gconsumo)
            cmd.Parameters.AddWithValue("@unidad", dts.gunidad)
            cmd.Parameters.AddWithValue("@consumo_total", dts.gconsumo_total)
            cmd.Parameters.AddWithValue("@costo_unitario", dts.gcosto_unitario)
            cmd.Parameters.AddWithValue("@costo_total", dts.gcosto_total)
            cmd.Parameters.AddWithValue("@id_cotizacion", dts.gid_cotizacion)

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
    Public Function mostrar_manoobra(ByVal dts As vManoObra) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_manoobra")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cotizacion", dts.gid_cotizacion)

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
    Public Function buscar_ManoObra_codigo(ByVal dts As vManoObra) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_ManoObra_codigo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cotizacion", dts.gid_cotizacion)

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
