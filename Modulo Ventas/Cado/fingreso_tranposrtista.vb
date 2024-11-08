Imports System.Data.SqlClient
Public Class fingreso_tranposrtista
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_ingre_transportista(ByVal dts As vingre_tranportistas) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_ingre_transportista")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@chofer", dts.gchofer)
            cmd.Parameters.AddWithValue("@marca", dts.gmarca)
            cmd.Parameters.AddWithValue("@placa", dts.gplaca)



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
