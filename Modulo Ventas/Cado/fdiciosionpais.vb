Imports System.Data.SqlClient
Public Class fdiciosionpais
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscar_cliente_comerial(ByVal dts As VPAIS_CO) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_GetAllZonasClientesComercial")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CODCLIENTE", dts.gruc)
            cmd.Parameters.AddWithValue("@nORDEN", dts.gorden)
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
    Public Function buscar_divicion_comerial(ByVal dts As VPAIS_CO) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_GetAllDivisionesClientesComercial")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CODCLIENTE", dts.gruc)
            cmd.Parameters.AddWithValue("@nORDEN", dts.gorden)
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
