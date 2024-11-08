Imports System.Data.SqlClient
Public Class fmoviles
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_moviles(ByVal dts As vmoviles) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_lineas")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@proveedor", dts.gproveedor)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@trabajador", dts.gtrabajador)
            cmd.Parameters.AddWithValue("@area", dts.garea)
            cmd.Parameters.AddWithValue("@mequipo", dts.gmequipo)
            cmd.Parameters.AddWithValue("@moequipo", dts.gmoequipo)
            cmd.Parameters.AddWithValue("@preequipo", dts.gpreequipo)
            cmd.Parameters.AddWithValue("@cuotas", dts.gcuotas)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            cmd.Parameters.AddWithValue("@plam", dts.gplam)
            cmd.Parameters.AddWithValue("@modalidad", dts.gmodalidad)
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)
            cmd.Parameters.AddWithValue("@fecha", dts.gfechas)


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
    Public Function mostrar_lineas(ByVal dts As vmoviles) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_linea")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@linea", dts.glinea)

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
    Public Function eliminar_movil(ByVal dts As vmoviles) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_linea")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@linea", dts.glinea)

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
