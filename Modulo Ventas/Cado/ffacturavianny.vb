Imports System.Data.SqlClient
Public Class ffacturavianny
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function mostrar_factura_notacredito(ByVal dts As vfacturasistema) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("nota_credito")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gtdoc)
            cmd.Parameters.AddWithValue("@serie", dts.gfatura)
            cmd.Parameters.AddWithValue("@correlativo", dts.gcorrelativo_fac)
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
    Public Function mostrar_factura_vianny_detalle(ByVal dts As vfacturasistema) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_factura_vianny_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)

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
    Public Function mostrar_factura_vianny_cabecera(ByVal dts As vfacturasistema) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_factura_vianny_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
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
    Public Function anular_factura(ByVal dts As vfacturasistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("anular_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@flag", dts.gflag)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
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
    Public Function actualizar_guiafactura(ByVal dts As vfacturasistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_guia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@sguia", dts.gsguia)
            cmd.Parameters.AddWithValue("@cguia", dts.gcguia)
            cmd.Parameters.AddWithValue("@sfactura", dts.gsfactura)
            cmd.Parameters.AddWithValue("@cfactura", dts.gcfactura)
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
    Public Function eiminar_factura_vianny(ByVal dts As vfacturasistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eiminar_factura_vianny")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
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
    Public Function insertar_sinonio(ByVal dts As vfacturasistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_sinonimo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@codigo_sin", dts.gcodigo_sin)
            cmd.Parameters.AddWithValue("@item_sin", dts.gitem_sin)
            cmd.Parameters.AddWithValue("@nomb_sin", dts.gnomb_sin)

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
    Public Function eliminar_sinonimo(ByVal dts As vfacturasistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_sinonimo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@codigo_sin", dts.gcodigo_sin)

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
    Public Function buscar_correlativo_venta(ByVal dts As vfacturavianny)
        Try
            conectare()
            cmd = New SqlCommand("correla_venta_almacen")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
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
    Public Function correlativo_fac_bol(ByVal dts As vfacturavianny)
        Try
            conectare()
            cmd = New SqlCommand("orrelativo_fac_bol")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fac").ToString()
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("EL NUMERODE SERIE NO EXISTE")
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function insertar_cabecera_venta(ByVal dts As vfacturasistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ingresar_cabecera_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
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
            cmd.Parameters.AddWithValue("@venta_sinigv_do", dts.gventa_sinigv_do)
            cmd.Parameters.AddWithValue("@igv_do", dts.gigv_do)
            cmd.Parameters.AddWithValue("@total_venta_do", dts.gtotal_venta_do)
            cmd.Parameters.AddWithValue("@glosa", dts.gglosa)
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@tipo_venta", dts.gtipo_venta)
            cmd.Parameters.AddWithValue("@afecto", dts.gafecto)
            cmd.Parameters.AddWithValue("@incluido", dts.gincluido)
            cmd.Parameters.AddWithValue("@glosafac", dts.gglosafac)
            cmd.Parameters.AddWithValue("@fechafin", dts.gfechafin)
            cmd.Parameters.AddWithValue("@stdoc", dts.gstdoc)
            cmd.Parameters.AddWithValue("@sfatura", dts.gsfatura)
            cmd.Parameters.AddWithValue("@scorrelativo_fac", dts.gscorrelativo_fac)
            cmd.Parameters.AddWithValue("@oper", dts.goper)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@pedido", dts.gpedido)

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
    Public Function insertar_detalle_venta(ByVal dts As vfacturadetallesistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ingresar_detalle_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
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
            cmd.Parameters.AddWithValue("@rubro", dts.grubro)
            cmd.Parameters.AddWithValue("@um", dts.gum)
            cmd.Parameters.AddWithValue("@grm", dts.ggrm)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@correlativo_fac", dts.gcorrelativo_fac)
            cmd.Parameters.AddWithValue("@fatura", dts.gfatura)
            cmd.Parameters.AddWithValue("@tido", dts.gtido)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@rubro_detallado", dts.grubrodetallado)
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
    Public Function eliminar_regventanegr(ByVal dts As vfacturasistema) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_ventan_sistema")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom)
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
End Class
