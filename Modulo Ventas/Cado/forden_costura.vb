Imports System.Data.SqlClient
Public Class forden_costura
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function COSTURA_CONTROLER() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("Asignacion_costura_CONTROLER")
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
    Public Function buscar_orden_costuraop(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_op_costura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gop)
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
    Public Function actualizar_cantidad(ByVal dts As vorden_costura) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ACTUALIZAR_CANTIDAD")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CANTIDAD", dts.gcantidad)
            cmd.Parameters.AddWithValue("@orden", dts.gorden)
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

    Public Function actualizar_costura(ByVal dts As vorden_costura) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ACTUALIZAR_estado_costura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@orden", dts.gorden)
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

    Public Function insertra_costura(ByVal dts As vorden_costura) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_costura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@orden_c", dts.gorden_c)
            cmd.Parameters.AddWithValue("@op_oc", dts.gop_oc)
            cmd.Parameters.AddWithValue("@ocorte", dts.gocorte)
            cmd.Parameters.AddWithValue("@prendasc", dts.gprendasc)
            cmd.Parameters.AddWithValue("@adelanto", dts.gadelanto)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
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
    Public Function buscar_costura(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_confeccion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gop)
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
    Public Function detalle_costura(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("detalle_costura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@orden", dts.gorden)
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
    Public Function fecha_produc_diaria(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("bucar_fecha_costura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gop)
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
    Public Function fecha_produc_diaria_10(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("bucar_fecha_costura_5")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gop)
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
    Public Function fecha_produc_diaria1(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("bucar_fecha_costura1")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@FECHA", dts.gfecha)
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
    Public Function ANALISIS_CONFECCION(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("ANALISIS_CONF_OP")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@op", dts.gop)
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
    Public Function BUSCAR_OP() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_OP")
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
    Public Function ANALISIS_CONFECCIONPO(ByVal dts As vorden_costura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("ANALISIS_CONF_PO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@PO", dts.gop)
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
