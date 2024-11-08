Imports System.Data.SqlClient
Public Class fnopedido
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_nota_pedido(ByVal dts As vnapedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_nota_pedido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@numero_pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@codigo_cli", dts.gcodigo_cli)
            cmd.Parameters.AddWithValue("@moneda", dts.gmoneda)
            cmd.Parameters.AddWithValue("@responsable", dts.gresponsable)
            cmd.Parameters.AddWithValue("@valor_comercial", dts.gvalor_comercial)
            cmd.Parameters.AddWithValue("@dir_despacho", dts.gdir_despacho)
            cmd.Parameters.AddWithValue("@forma_pago", dts.gforma_pago)
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@or_compra", dts.gor_compra)
            cmd.Parameters.AddWithValue("@f_enrtrega", dts.gf_enrtrega)
            cmd.Parameters.AddWithValue("@inmediato", dts.ginmediato)
            cmd.Parameters.AddWithValue("@observaciones", dts.gobservaciones)
            cmd.Parameters.AddWithValue("@sub_total", dts.gsub_total)
            cmd.Parameters.AddWithValue("@igv", dts.gigv)
            cmd.Parameters.AddWithValue("@total", dts.gtotal)

            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@pre_inclu_igv", dts.gpreinc)
            cmd.Parameters.AddWithValue("@estado_comer", dts.gescom)
            cmd.Parameters.AddWithValue("@estado_almacen", dts.gesalm)
            cmd.Parameters.AddWithValue("@estado_tranporte", dts.gesttrans)
            cmd.Parameters.AddWithValue("@cotizacion", dts.gcotizacion)
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)

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
    Public Function insertar_detalle_pedido(ByVal dts As cdetalle_pedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_detalle_pedido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@items", dts.gitems)
            cmd.Parameters.AddWithValue("@codigo_tela", dts.gcodigo_tela)
            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@densidad", dts.gdensidad)
            cmd.Parameters.AddWithValue("@ancho", dts.gancho)
            cmd.Parameters.AddWithValue("@um", dts.gum)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@precio_unitario", dts.gprecio_unitario)
            cmd.Parameters.AddWithValue("@igv", dts.gigv)
            cmd.Parameters.AddWithValue("@Total", dts.gTotal)
            cmd.Parameters.AddWithValue("@numero_pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@sub_Total", dts.gsubtotal)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@estado2", dts.gestado2)
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)
            cmd.Parameters.AddWithValue("@alm_num", dts.galm_num)
            cmd.Parameters.AddWithValue("@Rollos", dts.grollos)
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
    Public Function buscar_correo(ByVal dts As vnapedido)
        Try
            conectare()
            cmd = New SqlCommand("buscar_correo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vend", dts.gvendedor)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("correo").ToString()
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
    Public Function validar_pedido(ByVal dts As vnapedido)
        Try
            conectare()
            cmd = New SqlCommand("validar_notapedido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
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
    Public Function correlativo_npedido(ByVal dts As vnapedido)
        Try
            conectare()
            cmd = New SqlCommand("correlativo_npedido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function mostrar_pedidos(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_nota_pedido")
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
    Public Function buscar_nota_pedido_comercial(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_nota_pedido_comercial")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvendedor)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
            cmd.Parameters.AddWithValue("@PERIODO", dts.gperiodo)
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
    Public Function mostrar_pedidos1(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_nota_pedido1")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public saved As Boolean

    Public Function mostrar_pedidos2(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_nota_pedido2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
            If cmd.ExecuteNonQuery Then
                Dim dt As New DataTable
                Dim da As New SqlDataAdapter(cmd)
                da.Fill(dt)
                Return dt
                saved = True
            Else
                Return Nothing
                saved = False
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
            saved = False
        Finally
            desconectar()

        End Try
    End Function
    Public Function mostrar_pedidos_aprobado(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_nota_pedido_aprobado")
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
    Public Function mostrar_pedidos_aprobado1(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_nota_pedido_aprobado1")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function mostrar_pedidos_aprobado2(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_nota_pedido_aprobado2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function actualizar_ESTADO(ByVal dts As vnapedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ACTUALIZAR_ESTADO_NPEDIDO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ESTADO", dts.gestado)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function actualizar_ESTADO_comercial(ByVal dts As vnapedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ACTUALIZAR_ESTADO_NPEDIDO_comercial")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ESTADO", dts.gestado)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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

    Public Function buscar_pedido_cabecera(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("b_mpedido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function buscar_pedido_detalle(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("b_detalle_pedido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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

    Public Function validar_almacen(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("va_almacen")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function validar_despacho(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("va_despacho")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
            'cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
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

    Public Function eliminar_pedido(ByVal dts As vnapedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("elimar_notapedido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function ANULAR_pedido(ByVal dts As vnapedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ANULAR_NP")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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

    Public Function anular_pedidos_automaticos() As Boolean
        Try
            conectare()
            cmd = New SqlCommand("anular_pedidos_automaticos")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
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

    Public Function actualizar_estado_almacen(ByVal dts As vnapedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_estadoalmacen")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
    Public Function actualizar_estado_despacho(ByVal dts As vnapedido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_estado_despacho")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@pedido", dts.gnumero_pedido)
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
    Public Function validar_almacen2(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("va_almacen2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
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
    Public Function validar_despacho2(ByVal dts As vnapedido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("va_despacho2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gempresa)
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
