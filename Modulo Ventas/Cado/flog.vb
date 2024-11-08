Imports System.Data.SqlClient
Public Class flog
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_log(ByVal dts As vlog) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_LOG")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@modulo", dts.gmodulo)
            cmd.Parameters.AddWithValue("@cuser", dts.gcuser)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@pc", dts.gpc)
            cmd.Parameters.AddWithValue("@accion", dts.gaccion)
            cmd.Parameters.AddWithValue("@detalle", dts.gdetalle)
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
End Class
