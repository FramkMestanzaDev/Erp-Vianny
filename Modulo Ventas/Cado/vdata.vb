Imports System.Data.SqlClient
Public Class vdata
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscar_partida(ByVal dts As fadat) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_PARTIDA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
            If cmd.ExecuteNonQuery Then
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox("Partida no Existe")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function mostrar_partida_rollo(ByVal dts As fadat) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_partida_rollo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)

            If cmd.ExecuteNonQuery Then
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox("Partida no Existe")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function

    Public Function actulizar_calidad(ByVal dts As fplanea) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("altualizar_fecha")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@rollo", dts.grollo)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
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

    Public Function actulizar_rollo(ByVal dts As fplanea) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("altualizar_rollo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@rollo", dts.grollo)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@peso", dts.gpeso)
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
    Public Function actulizar_calidad_TELA(ByVal dts As fadat) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("altualizar_calidad_tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@rollo", dts.grollo)
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
