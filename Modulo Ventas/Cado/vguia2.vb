Imports System.Data.SqlClient
Public Class vguia2
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function consultar_detalle_guia(ByVal dts As fguia2) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_guia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)

            If cmd.ExecuteNonQuery Then
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox("boleta no existe")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function consultar_cabecera_guia(ByVal dts As fguia2) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_guia_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)

            If cmd.ExecuteNonQuery Then
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox("Guia no existe")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
End Class
