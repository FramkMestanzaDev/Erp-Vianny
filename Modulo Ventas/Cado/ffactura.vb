Imports System.Data.SqlClient
Public Class ffactura
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscar_cobranzas_vendedor(ByVal dts As vfactura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("cobranzas_mensuales_vendedor]")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)

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
    Public Function buscar_facturas_anuales(ByVal dts As vfactura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("ventas_anulaes_mes")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function buscar_facturas_cobranzas(ByVal dts As vfactura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("cobranzas_mensuales")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function buscar_facturas() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_factura")
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
    Public Function actualizar_factura(ByVal dts As vfactura) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)

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
    Public Function actualizar_factura_MONTO(ByVal dts As vfactura) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("MONTO_PARCIAL")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@MONTO", dts.gMONTO)
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
    Public Function insertar_ventas_fac_reg(ByVal dts As vdet_regis_fact) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_ventas_fac_reg")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@comprobante", dts.gcomprobante)
            cmd.Parameters.AddWithValue("@cliente", dts.gcliente)
            cmd.Parameters.AddWithValue("@valor_venta", dts.gvalor_venta)
            cmd.Parameters.AddWithValue("@igv", dts.gigv)
            cmd.Parameters.AddWithValue("@total", dts.gtotal)
            cmd.Parameters.AddWithValue("@pagado", dts.gpagado)
            cmd.Parameters.AddWithValue("@parcial", dts.gparcial)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
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
    Public Function buscar_facturas_anuales_NEGRAS() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("cobranzas_mensuales_NEGRAS")
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
    Public Function buscar_facturas_anuales_NEGRAS_ventas() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("ventas_anulaes_mes_negras")
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
    Public Function comisiones_vendedor(ByVal dts As vfactura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("COMISIONES_vendedor_detallado")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@PERIODO", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function comisiones_vendedor_edicion(ByVal dts As vfactura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("COMISIONES_vendedor_edicion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@PERIODO", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function comisiones_vendedor_detallado(ByVal dts As vfactura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("COMISIONES_vendedor_detallado_total")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@PERIODO", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function comisiones_edicion_total(ByVal dts As vfactura) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("COMISIONES_edicion_total")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@cperiodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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


    Public Function mostrar_estado_facbol(ByVal dts As vfactura)
        Try
            conectare()
            cmd = New SqlCommand("buscar_estado_pago")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@ndoc", dts.gndoc)
            cmd.Parameters.AddWithValue("@CLIENTE", dts.gCLIENTE)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("SALDO").ToString()
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function mostrar_estado_facbol2(ByVal dts As vfactura)
        Try
            conectare()
            cmd = New SqlCommand("buscar_estado_pago2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@ndoc", dts.gndoc)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("LET").ToString()
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function mostrar_estado_facbol3(ByVal dts As vfactura)
        Try
            conectare()
            cmd = New SqlCommand("buscar_estado_canjeado")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@ndoc", dts.gndoc)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("CAN").ToString()
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
End Class
