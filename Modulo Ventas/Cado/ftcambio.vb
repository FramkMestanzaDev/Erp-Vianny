Imports System.Data.SqlClient
Public Class ftcambio
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function mostrar_tipo_cambio(ByVal dts As vtcambio)
        Try
            conectare()
            cmd = New SqlCommand("tipo_cambio")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("tcamv_23i").ToString()
            Else
                Return "0.0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
End Class
