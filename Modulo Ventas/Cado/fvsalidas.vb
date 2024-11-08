Imports System.Data.SqlClient
Public Class fvsalidas

    Inherits conexion
        Dim cmd As New SqlCommand
    Public Function Valorizar_Salidas(ByVal dts As vvsalidas) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("Valorizar_Salidas")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.CommandTimeout = 220
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@ano", dts.gano)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)
            If dts.gmotivos.ToString.Length = 0 Then
                cmd.Parameters.AddWithValue("@motivos", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@motivos", dts.gmotivos)
            End If
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
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
    Public Function listar_codigos(ByVal dts As vvsalidas) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("listar_codigos")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)
            cmd.Parameters.AddWithValue("@ano", dts.gano)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@mes", dts.gmes)

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
