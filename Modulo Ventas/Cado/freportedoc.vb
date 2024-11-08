Imports System.Data.SqlClient
Public Class freportedoc
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function reportedoc1(ByVal dts As vreportedoc) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("reporte1")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fechaini", dts.gfechaini)
            cmd.Parameters.AddWithValue("@fechafin", dts.gfechafin)
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
    Public Function reportedoc2(ByVal dts As vreportedoc) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("REPORTE2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fechaini", dts.gfechaini)
            cmd.Parameters.AddWithValue("@fechafin", dts.gfechafin)
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
    Public Function reportedoc3(ByVal dts As vreportedoc) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_plog")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fecha", dts.gfechaini)

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
    Public Function reportedoc4(ByVal dts As vreportedoc) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_plog2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)

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
    Public Function reportedoc5() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_plog3")
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
    Public Function reportedoc6(ByVal dts As vreportedoc) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_plog2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gfactura2)

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
    Public Function reporte_tesoreria() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_tesoreria")
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
    Public Function reporte_tesoreria2() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_tesoreria2")
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
    Public Function actualizar_ESTADO(ByVal dts As vreportedoc) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("atualizar_registro_tesoreria")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gestado)

            If cmd.ExecuteNonQuery Then
                Return True
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
