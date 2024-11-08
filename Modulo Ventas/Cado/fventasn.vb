Imports System.Data.SqlClient
Public Class fventasn
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertarventasn(ByVal dts As vventas_n) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_ventas_negras")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@salida", dts.gsalida)
            cmd.Parameters.AddWithValue("@comprobante", dts.gcomprobante)
            cmd.Parameters.AddWithValue("@cliente", dts.gcliente)
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@fpago", dts.gfpago)
            cmd.Parameters.AddWithValue("@valor_venta", dts.gvalor_venta)
            cmd.Parameters.AddWithValue("@igv", dts.gigv)
            cmd.Parameters.AddWithValue("@total", dts.gtotal)
            cmd.Parameters.AddWithValue("@pagado", dts.gpagado)
            cmd.Parameters.AddWithValue("@parcial", dts.gparcial)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
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
    Public Function buscar_ventasNN() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_VENTASNN")
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
    Public Function buscar_ventasNN2(ByVal dts As vventas_n) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_RVENTANN")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@clientes", dts.gcliente)

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
    Public Function BUSCAR_DETALLE_LIBRE(ByVal dts As vventas_n) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_DETALLE_LIBRE")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@PERIODO", dts.gperiodo)
            cmd.Parameters.AddWithValue("@TIPO", dts.gTIPO)
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
    Public Function verificar_guia(ByVal dts As vventas_n)
        Try
            conectare()
            cmd = New SqlCommand("buscar_guia_nsa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@guia_nsa", dts.gguia_nsa)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("CANT").ToString()
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
    Public Function verificar_guia_vianny(ByVal dts As vventas_n)
        Try
            conectare()
            cmd = New SqlCommand("buscar_guia_nsa_vianny")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@guia_nsa", dts.gguia_nsa)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("CANT").ToString()
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
    Public Function buscartdoc(ByVal dts As vventas_n)
        Try
            conectare()
            cmd = New SqlCommand("buscar_tdoc")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@guia", dts.gguia_nsa)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("tdoc_10").ToString()
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
End Class
