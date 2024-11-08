Imports System.Data.SqlClient
Public Class combiancines
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function mostrar_colores(ByVal dts As vproducot) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("combinaciones")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
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
