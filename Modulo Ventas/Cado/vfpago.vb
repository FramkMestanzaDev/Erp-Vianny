Imports System.Data.SqlClient
Public Class vfpago
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function forma_pago() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("fo_pago")
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
