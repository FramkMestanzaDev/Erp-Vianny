Imports System.Data.SqlClient
Public Class t_venta
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscatipo_ventar_op() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("Mostar_Tipo_Venta")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            If cmd.ExecuteNonQuery Then
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox("No Existe Informacion")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function


    Public Function buscar_cabecera_venta(ByVal dts As vvventas) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_factura_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@docu", dts.gdocumento)
            cmd.Parameters.AddWithValue("@sfac", dts.gserie)
            cmd.Parameters.AddWithValue("@nfac", dts.gcorrela)
            If cmd.ExecuteNonQuery Then
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox("No Existe Informacion")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function

    Public Function buscar_detalle_venta(ByVal dts As vvventas) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@docu", dts.gdocumento)
            cmd.Parameters.AddWithValue("@sfac", dts.gserie)
            cmd.Parameters.AddWithValue("@nfac", dts.gcorrela)
            If cmd.ExecuteNonQuery Then
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Return dt
            Else
                Return Nothing

            End If
        Catch ex As Exception
            MsgBox("No Existe Informacion")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
End Class
