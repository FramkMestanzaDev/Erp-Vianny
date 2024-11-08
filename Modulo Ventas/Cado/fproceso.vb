Imports System.Data.SqlClient
Public Class fproceso

    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function bucar_proceso(ByVal dts As vproceso) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("bucar_proceso")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
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
End Class
