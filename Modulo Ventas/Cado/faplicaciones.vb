Imports System.Data.SqlClient
Public Class faplicaciones
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_aplicaciones(ByVal dts As Aplicaciones) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_AplicacioneS")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@merma", dts.gmerma)
            cmd.Parameters.AddWithValue("@consumo", dts.gconsumo)
            cmd.Parameters.AddWithValue("@unidad", dts.gunidad)
            cmd.Parameters.AddWithValue("@consumo_tot", dts.gconsumo_tot)
            cmd.Parameters.AddWithValue("@cunitario", dts.gcosto_unitario)
            cmd.Parameters.AddWithValue("@costo_total", dts.gcosto_total)
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
    Public Function mostrar_aplicaciones(ByVal dts As Aplicaciones) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_aplicaciones")
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
    Public Function buscar_Aplicaciones_codigo(ByVal dts As Aplicaciones) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_Aplicaciones_codigo")
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
