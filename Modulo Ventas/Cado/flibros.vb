Imports System.Data.SqlClient
Public Class flibros
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function Reporte_Ventas(ByVal dts As vlibros) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteRegistroVentasNuevo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CPERIODO", dts.gCPERIODO)
            cmd.Parameters.AddWithValue("@fecha_ini", dts.gfecha_ini)
            cmd.Parameters.AddWithValue("@fecha_FIN", dts.gfecha_FIN)
            cmd.Parameters.AddWithValue("@Moneda", 1)
            cmd.Parameters.AddWithValue("@Tipo_Venta", "NULL")
            cmd.Parameters.AddWithValue("@Tienda", "NULL")
            cmd.Parameters.AddWithValue("@nQuiebre", 3)
            cmd.Parameters.AddWithValue("@NORDEN", 1)
            cmd.Parameters.AddWithValue("@ntipoAlmacen", 1)

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
    Public Function Reporte_Compras(ByVal dts As vlibros) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteRegistroComprasPLE")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CPERIODO", dts.gCPERIODO)
            cmd.Parameters.AddWithValue("@fecha_ini", dts.gfecha_ini)
            cmd.Parameters.AddWithValue("@fecha_FIN", dts.gfecha_FIN)
            cmd.Parameters.AddWithValue("@Tipo_Compra", DBNull.Value)
            cmd.Parameters.AddWithValue("@Moneda", 1)
            cmd.Parameters.AddWithValue("@NORDEN", 1)
            cmd.Parameters.AddWithValue("@nQuiebre", 1)


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
    Public Function Libro_Diario(ByVal dts As vlibros) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteLibroDiarioNuevo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCIA", dts.gccia)
            cmd.Parameters.AddWithValue("@CPERIODO", dts.gCPERIODO)
            cmd.Parameters.AddWithValue("@CMES", dts.gmes)
            cmd.Parameters.AddWithValue("@nMoneda", 1)



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
    Public Function Libro_Diario5_3(ByVal dts As vlibros) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteLibroDiarioNuevo_5_3")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCIA", dts.gccia)
            cmd.Parameters.AddWithValue("@CPERIODO", dts.gCPERIODO)
            cmd.Parameters.AddWithValue("@CMES", dts.gmes)
            cmd.Parameters.AddWithValue("@nMoneda", 1)



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
