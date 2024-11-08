Imports System.Data.SqlClient
Public Class fdireccion
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_direccion(ByVal dts As vdetdireccion) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_detalle_dire")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@direccion", dts.gdirecion)


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

    Public Function buscar_direc_despacc(ByVal dts As vdetdireccion) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_dir_des")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
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
