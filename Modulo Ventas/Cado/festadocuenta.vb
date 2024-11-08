Imports System.Data.SqlClient
Public Class festadocuenta
    Inherits conexion
    Dim cmd As New SqlCommand
    'Public Function buscar_estado_cuenta(ByVal dts As vestadocuenta) As DataTable
    '    Try
    '        conectare()
    '        cmd = New SqlCommand("UNIR_PROCEDIMIENTO3")
    '        cmd.CommandType = CommandType.StoredProcedure
    '        cmd.Connection = cnn

    '        cmd.Parameters.AddWithValue("@fecha", dts.gFECHA)
    '        cmd.Parameters.AddWithValue("@ruc", dts.gruc)
    '        If cmd.ExecuteNonQuery Then
    '            Dim dt As New DataTable
    '            Dim da As New SqlDataAdapter(cmd)
    '            da.Fill(dt)
    '            Return dt
    '        Else
    '            Return Nothing

    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message)

    '    Finally
    '        desconectar()

    '    End Try

    'End Function
    Public Function buscar_estado_cuenta(ByVal dts As vestadocuenta) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("UNIR_PROCEDIMIENTO3")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.CommandTimeout = 120
            cmd.Parameters.AddWithValue("@fecha", dts.gFECHA)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
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
    Public Function buscar_estado_cuenta2() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("UNIR_PROCEDIMIENTO4")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.CommandTimeout = 120

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
    Public Function buscar_estado_cuenta1(ByVal dts As vestadocuenta) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_estado_cuenta1")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)

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
