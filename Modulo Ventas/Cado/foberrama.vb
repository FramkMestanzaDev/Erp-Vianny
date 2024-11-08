Imports System.Data.SqlClient
Public Class foberrama
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function INSERTAR_OBSERVACIONES_RAMA(ByVal dts As vrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_OBSERVACIONES_RAMA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@PROGRAMA", dts.gPROGRAMA)
            cmd.Parameters.AddWithValue("@FECHA", dts.gfecha)
            cmd.Parameters.AddWithValue("@HORA", dts.ghora)
            cmd.Parameters.AddWithValue("@MOTIVO", dts.gMOTIVO)
            cmd.Parameters.AddWithValue("@COMENTARIO", dts.gCOMENTARIO)


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
