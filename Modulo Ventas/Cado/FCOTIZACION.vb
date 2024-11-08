Imports System.Data.SqlClient
Public Class FCOTIZACION
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function CORRELATIVO_COTIZACION()
        Try
            conectare()
            cmd = New SqlCommand("CORRELA_COTIZACION")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("VAL").ToString()
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

    Public Function insertar_cotizacion(ByVal dts As VCOTIZACION) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_HCOTIZACION")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@id_cotizacion", dts.gid_cotizacion)
            cmd.Parameters.AddWithValue("@cliente", dts.gcliente)
            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@estilo", dts.gestilo)
            cmd.Parameters.AddWithValue("@rango_tallas", dts.grango_tallas)
            cmd.Parameters.AddWithValue("@ti_cambio", dts.gti_cambio)
            cmd.Parameters.AddWithValue("@moneda", dts.gmoneda)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@gasto_logistico", dts.ggasto_logistico)
            cmd.Parameters.AddWithValue("@gasto_operativo", dts.ggasto_operativo)
            cmd.Parameters.AddWithValue("@Gastos_administrativos", dts.ggasto_administrativo)
            cmd.Parameters.AddWithValue("@gastos_financieros", dts.ggasto_financiero)
            cmd.Parameters.AddWithValue("@gastos_venta", dts.ggasto_venta)
            cmd.Parameters.AddWithValue("@costo_producto", dts.gcosto_producto)
            cmd.Parameters.AddWithValue("@margen_utilidad", dts.gmargen_utilidad)
            cmd.Parameters.AddWithValue("@valor_venta", dts.gvalor_venta)
            cmd.Parameters.AddWithValue("@igv", dts.gigv)
            cmd.Parameters.AddWithValue("@precio_venta", dts.gprecio_venta)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@imagen", dts.gimagen)
            cmd.Parameters.AddWithValue("@costot_tela", dts.gcostot_tela)
            cmd.Parameters.AddWithValue("@costot_AviosC", dts.gcostot_AviosC)
            cmd.Parameters.AddWithValue("@costot_AviosA", dts.gcostot_AviosA)
            cmd.Parameters.AddWithValue("@costot_Lavado", dts.gcostot_Lavado)
            cmd.Parameters.AddWithValue("@costot_Aplicaciones", dts.gcostot_Aplicaciones)
            cmd.Parameters.AddWithValue("@costot_MObra", dts.gcostot_MObra)
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@margen_utilidad_moneda", dts.gmargen_utilidad_moneda)
            cmd.Parameters.AddWithValue("@od", dts.god)
            cmd.Parameters.AddWithValue("@od_version", dts.god_version)
            cmd.Parameters.AddWithValue("@Pm", dts.gPm)
            cmd.Parameters.AddWithValue("@Tela", dts.gTela)
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
    Public Function eliminar_cotizacion(ByVal dts As VCOTIZACION) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_informacoion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cotizacion", dts.gid_cotizacion)
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
    Public Function mostrar_COTIZACION(ByVal dts As VCOTIZACION)
        Try
            conectare()
            cmd = New SqlCommand("buscar_COTIZACION")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cotizacion", dts.gid_cotizacion)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return Convert.ToInt32(dr("Cotizacion"))
            Else
                Return Convert.ToInt32(dr("Cotizacion"))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
        Finally
            desconectar()

        End Try
    End Function

    Public Function mostrar_COTIZACION2(ByVal dts As VCOTIZACION) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("hoja_cotizacion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cotizacion", dts.gid_cotizacion)

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
    Public Function mostrar_COTIZACION3() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("baucar_cotiza")
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
    Public Function mostrar_COTIZACION4() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("baucar_cotiza2")
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
    Public Function mostrar_costos(ByVal dts As VCOTIZACION) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_costp")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@tipo", dts.gtipo)
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
    Public Function actualizar_ESTADO(ByVal dts As VCOTIZACION) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ACTUALIZAR_ESTADO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@COTIZACION", dts.gid_cotizacion)
            cmd.Parameters.AddWithValue("@ESTADO", dts.gestado)
            If cmd.ExecuteNonQuery Then
                Return True
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
    Public Function mostrar_TODASCOTIZACIONES() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("TODAS_COTIZACIONES")
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
    Public Function reportedoc1(ByVal dts As VCOTIZACION) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("reporte1_cotizacion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fechaini", dts.gfinicial)
            cmd.Parameters.AddWithValue("@fechafin", dts.gffinal)
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
    Public Function reportedoc2(ByVal dts As VCOTIZACION) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("REPORTE2_cotizacion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@fechaini", dts.gfinicial)
            cmd.Parameters.AddWithValue("@fechafin", dts.gffinal)
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
    Public Function ganancias(ByVal dts As VCOTIZACION) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("ganancias_vendedor")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)

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
    Public Function buscar_vendedor_cotizacion(ByVal dts As VCOTIZACION) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("baucar_cotiza3")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)

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
    Public Function UltimoPrecioOc(ByVal dts As VCOTIZACION)
        Try
            conectare()
            cmd = New SqlCommand("Sp_ListadoUltimoPrecios")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCIA", dts.gempresa)
            cmd.Parameters.AddWithValue("@PROVEEDOR", dts.gproveedor)
            cmd.Parameters.AddWithValue("@ARTICULO", dts.garticulo)
            cmd.Parameters.AddWithValue("@FLIMITE", dts.gflimite)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("PUSoles").ToString()
            Else
                Return "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
End Class
