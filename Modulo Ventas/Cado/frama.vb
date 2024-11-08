Imports System.Data.SqlClient
Class frama
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function reporte_ram() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("reporte_ram")
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
    Public Function seguimiento_rama(ByVal dts As vrama) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("seguimiento_rama")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
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
    Public Function buscar_partida(ByVal dts As vrama) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_partida_rama")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function buscar_codigo_rama(ByVal dts As vrama) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_rama")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function actualizar_RAMA(ByVal dts As vrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actilizar_pasado")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@HORA", dts.ghora)
            cmd.Parameters.AddWithValue("@CODIGO", dts.gcodigo)
            cmd.Parameters.AddWithValue("@id", dts.gid)

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
    Public Function actualizar_RAMA_fin(ByVal dts As vrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_rama_fin")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@HORA", dts.ghora)
            cmd.Parameters.AddWithValue("@CODIGO", dts.gcodigo)
            cmd.Parameters.AddWithValue("@id", dts.gid)

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
    Public Function insertar_rama(ByVal dts As vrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_rama")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@orden", dts.gorden)
            cmd.Parameters.AddWithValue("@po", dts.gpo)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo)
            cmd.Parameters.AddWithValue("@rollos", dts.grollos)
            cmd.Parameters.AddWithValue("@peso", dts.gpeso)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@ancho", dts.gancho)
            cmd.Parameters.AddWithValue("@densidad", dts.gdensidad)
            cmd.Parameters.AddWithValue("@flag", dts.gflag)
            cmd.Parameters.AddWithValue("@peso_total", dts.gpeso)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@fecha_rama", dts.gfecha_rama)
            cmd.Parameters.AddWithValue("@hora", dts.ghora)
            cmd.Parameters.AddWithValue("@pasado", dts.gpasado)
            cmd.Parameters.AddWithValue("@hora2", dts.ghora2)
            cmd.Parameters.AddWithValue("@anchor", dts.ganchor)
            cmd.Parameters.AddWithValue("@densidadr", dts.gdensidadr)
            cmd.Parameters.AddWithValue("@velocidad", dts.gvelocidad)
            cmd.Parameters.AddWithValue("@temperatura", dts.gtemperatura)
            cmd.Parameters.AddWithValue("@igsa", "0")
            cmd.Parameters.AddWithValue("@sobrealimentacion", dts.gsalimentacion)
            cmd.Parameters.AddWithValue("@densidad_reposa", "")
            cmd.Parameters.AddWithValue("@maquina", dts.gmaquina)
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
    Public Function eliminar_rama(ByVal dts As vrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_rama")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
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
    Public Function correlativo_rama(ByVal dts As vrama)
        Try
            conectare()
            cmd = New SqlCommand("correlativo_rama")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("codigo").ToString()
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
    Public Function ANULAR_rama(ByVal dts As vrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("anular_rama")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
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
End Class
