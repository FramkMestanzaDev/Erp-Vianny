Imports System.Data.SqlClient
Public Class fgu_comprobante
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_cabecera_venta(ByVal dts As vgu_comprobante) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ingresar_cabecera_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
            cmd.Parameters.AddWithValue("@tdoc", dts.gtdoc)
            cmd.Parameters.AddWithValue("@fatura", dts.gfatura)
            cmd.Parameters.AddWithValue("@correlativo_fac", dts.gcorrelativo_fac)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@condicion_pago", dts.gcondicion_pago)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@razon_social", dts.grazon_social)
            cmd.Parameters.AddWithValue("@moneda", dts.gmoneda)
            cmd.Parameters.AddWithValue("@cambio", dts.gcambio)
            cmd.Parameters.AddWithValue("@venta_total", dts.gventa_total)
            cmd.Parameters.AddWithValue("@venta_sinigv", dts.gventa_sinigv)
            cmd.Parameters.AddWithValue("@igv", dts.gigv)
            cmd.Parameters.AddWithValue("@total_venta", dts.gtotal_venta)
            cmd.Parameters.AddWithValue("@venta_total_do", dts.gtotal_venta_do)
            cmd.Parameters.AddWithValue("@venta_sinigv_do", dts.gventa_sinigv_do)
            cmd.Parameters.AddWithValue("@igv_do", dts.gigv_do)
            cmd.Parameters.AddWithValue("@total_venta_do", dts.gtotal_venta_do)
            cmd.Parameters.AddWithValue("@glosa", dts.gglosa)
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@tipo_venta", dts.gtipo_venta)
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
    Public Function buscar_coprobante(ByVal dts As vnia)
        Try
            conectare()
            cmd = New SqlCommand("correla_venta")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ncom").ToString()
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
    Public Function buscar_coprobante_nsa(ByVal dts As vnia)
        Try
            conectare()
            cmd = New SqlCommand("correla_nsa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen1)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ncom").ToString()
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

    Public Function insertar_detalle_venta(ByVal dts As vdetaleventa) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ingresar_detalle_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom ", dts.gncom1)
            cmd.Parameters.AddWithValue("@items", dts.gitems)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@sinonimo", dts.gsinonimo)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@precio_unitario", dts.gprecio_unitario)
            cmd.Parameters.AddWithValue("@venta_total", dts.gventa_total)
            cmd.Parameters.AddWithValue("@valor_venta", dts.gvalor_venta)
            cmd.Parameters.AddWithValue("@igv", dts.gigv)
            cmd.Parameters.AddWithValue("@total", dts.gtotal)
            cmd.Parameters.AddWithValue("@precio_unitario2", dts.gprecio_unitario2)
            cmd.Parameters.AddWithValue("@venta_total2", dts.gventa_total2)
            cmd.Parameters.AddWithValue("@valor_venta2", dts.gvalor_venta2)
            cmd.Parameters.AddWithValue("@igv2", dts.gigv2)
            cmd.Parameters.AddWithValue("@total2", dts.gtotal2)
            cmd.Parameters.AddWithValue("@op", dts.gop)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)

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

    Public Function insertar_sinonimo(ByVal dts As vsinonimo) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_sinonimo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@iten", dts.gitem)
            cmd.Parameters.AddWithValue("@descrip", dts.gdescrip)

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
    Public Function actualizar_sinonimo(ByVal dts As partida) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_sinonimo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@nombre", dts.gnombre)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)


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

    Public Function buscar_sinonimo(ByVal dts As partida) As String
        Try
            conectare()
            cmd = New SqlCommand("buscar_sininimo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("dato").ToString()
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
