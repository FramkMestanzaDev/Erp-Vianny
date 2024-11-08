Imports System.Data.SqlClient
Public Class ftenido
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_tenido_cab(ByVal dts As vtenido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_teñido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ntenido", dts.gntenido)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@metraje", dts.gmetraje)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@nhoja", dts.gnhoja)
            cmd.Parameters.AddWithValue("@estdo", dts.gestado)

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
    Public Function insertar_tenido_deta(ByVal dts As vtenido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_detalle_tenido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ntenido", dts.gntenido)
            cmd.Parameters.AddWithValue("@hora", dts.ghora)
            cmd.Parameters.AddWithValue("@ph", dts.gph)
            cmd.Parameters.AddWithValue("@mv", dts.gmv)
            cmd.Parameters.AddWithValue("@indigo", dts.gindigo)
            cmd.Parameters.AddWithValue("@hidro", dts.ghidro)
            cmd.Parameters.AddWithValue("@be", dts.gbe)
            cmd.Parameters.AddWithValue("@veloc", dts.gveloc)
            cmd.Parameters.AddWithValue("@tinas", dts.gtinas)
            cmd.Parameters.AddWithValue("@ds_colorant", dts.gds_colorant)
            cmd.Parameters.AddWithValue("@ds_hidrosuf", dts.gds_hidrosuf)
            cmd.Parameters.AddWithValue("@ds_soda", dts.gds_soda)
            cmd.Parameters.AddWithValue("@ds_pquimic", dts.gds_pquimic)
            cmd.Parameters.AddWithValue("@ds_neutral", dts.gds_neutral)
            cmd.Parameters.AddWithValue("@ds_fijad", dts.gds_fijad)
            cmd.Parameters.AddWithValue("@ph_neutral", dts.gph_neutral)
            cmd.Parameters.AddWithValue("@ph_fijad", dts.gph_fijad)
            cmd.Parameters.AddWithValue("@ph_suaviz", dts.gph_suaviz)
            cmd.Parameters.AddWithValue("@sol", dts.gsol)
            cmd.Parameters.AddWithValue("@observaciones", dts.gobservaciones)
            cmd.Parameters.AddWithValue("@estado", dts.gestado1)
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
    Public Function GRAFICO_TENIDO() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("GRAFICO_TENIDO")
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
    Public Function buscar_cabecera(ByVal dts As vtenido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("GRAFICO_TENIDO_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@tenido", dts.gntenido)


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
    Public Function buscar_detalle(ByVal dts As vtenido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("GRAFICO_TENIDO_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@tenido", dts.gntenido)

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
    Public Function eliminar_cabecer_tenido(ByVal dts As vtenido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_TENIDO_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@tenido", dts.gntenido)

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
    Public Function actualizar_estado(ByVal dts As vtenido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_estado_tenido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@tenido", dts.gntenido)

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
