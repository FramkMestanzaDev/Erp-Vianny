Imports System.Data.SqlClient
Public Class ftela
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_tela(ByVal dts As vTela) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_Tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@merma", dts.gmerma)
            cmd.Parameters.AddWithValue("@consumo", dts.gconsumo)
            cmd.Parameters.AddWithValue("@unidad", dts.gunidad)
            cmd.Parameters.AddWithValue("@consumo_total", dts.gconsumo_total)
            cmd.Parameters.AddWithValue("@cost_unitario", dts.gccosto_unitario)
            cmd.Parameters.AddWithValue("@cost_total", dts.gcosto_total)
            cmd.Parameters.AddWithValue("@items", dts.gitems)
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
    Public Function mostrar_tela(ByVal dts As vTela) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_tela")
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
        End

    End Function
    Public Function buscar_tela_codigo(ByVal dts As vTela) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_tela_codigo")
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
        End
    End Function
End Class
