Imports System.Data.SqlClient
Public Class fanalisis
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscar_informacion(ByVal dts As vanalisis) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("unir_procedimeinto_5")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fecini", dts.gfini)
            cmd.Parameters.AddWithValue("@fechfin", dts.gffin)
            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
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
