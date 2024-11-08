Imports System.Data.SqlClient
Public Class ffamilia
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function mostrar_familia() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_FAMILIA")
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
End Class
