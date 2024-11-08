Imports System.Data.SqlClient
Public Class fcliente
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscar_codigo_almacen(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_CODIGO")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
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
    Public Function buscar_codigo_almacen2(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_CODIGO_10")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
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
    Public Function buscar_cliente_COMERCIAL(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CLIENTE_COMERCIAL")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@VENDEDOR", dts.gVendedor)
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
    Public Function buscar_VENDEDOR_CLIENTE12(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CLIENTE_VENDEDOR")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@VENDEDOR", dts.gVendedor)
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
    Public Function CORRELATIVO_CLIENTE()
        Try
            conectare()
            cmd = New SqlCommand("correlativo_codigo")
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
    Public Function insertar_validar_partida(ByVal dts As vvalidarpartida) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("INSERTAR_VALIDAR_PARTIDAS")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo)
            cmd.Parameters.AddWithValue("@cantrollos", dts.gcantrollos)
            cmd.Parameters.AddWithValue("@kgstejidos", dts.gkgsobtenidos)
            cmd.Parameters.AddWithValue("@merma", dts.gmerma)
            cmd.Parameters.AddWithValue("@rollospesados", dts.grollospesados)
            cmd.Parameters.AddWithValue("@kgsobtenidos", dts.gkgsobtenidos)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
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
    Public Function insertar_cliente(ByVal dts As vcliente) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_cliente")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@T_persona", dts.gT_persona)
            cmd.Parameters.AddWithValue("@r_social", dts.gr_social)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@D_fiscal", dts.gD_fiscal)
            cmd.Parameters.AddWithValue("@email", dts.gemail)
            cmd.Parameters.AddWithValue("@telefono", dts.gtelefono)
            cmd.Parameters.AddWithValue("@email_fac", dts.gemail_fac)
            cmd.Parameters.AddWithValue("@Vendedor", dts.gVendedor)
            cmd.Parameters.AddWithValue("@origen", dts.gorigen)
            cmd.Parameters.AddWithValue("@pais", dts.gpais)
            cmd.Parameters.AddWithValue("@departamento", dts.gdepartamento)
            cmd.Parameters.AddWithValue("@provincia", dts.gprovincia)
            cmd.Parameters.AddWithValue("@distrito", dts.gdistrito)
            cmd.Parameters.AddWithValue("@nombres", dts.gnombres)
            cmd.Parameters.AddWithValue("@apaterno", dts.gapaterno)
            cmd.Parameters.AddWithValue("@amaterno", dts.gamaterno)
            cmd.Parameters.AddWithValue("@t_doc", dts.gt_doc)
            cmd.Parameters.AddWithValue("@contacto", dts.gcontacto)
            cmd.Parameters.AddWithValue("@COD_CLI", dts.gCOD_CLI)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@lcredito", dts.glcredito)
            cmd.Parameters.AddWithValue("@U_NEGOCIO", dts.gU_NEGOCIO)
            cmd.Parameters.AddWithValue("@fcom", dts.gfcom)
            cmd.Parameters.AddWithValue("@fcom_fin", dts.gfcom_fin)
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
    Public Function insertar_cliente_vianny(ByVal dts As vclientevianny) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_cliente_vianny")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@nombre", dts.gnombre)
            cmd.Parameters.AddWithValue("@tdoc", dts.gtdoc)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@direccion", dts.gdireccion)
            cmd.Parameters.AddWithValue("@origen", dts.gorigen)
            cmd.Parameters.AddWithValue("@pais", dts.gpais)
            cmd.Parameters.AddWithValue("@email_fac", dts.gemail_fac)
            cmd.Parameters.AddWithValue("@ubigeo", dts.gubigeo)
            cmd.Parameters.AddWithValue("@tiper", dts.gtiper)
            cmd.Parameters.AddWithValue("@nomf", dts.gnomf)
            cmd.Parameters.AddWithValue("@amaterno", dts.gamaterno)
            cmd.Parameters.AddWithValue("@apaterno", dts.gapaterno)


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

    Public Function actualizar_cliente(ByVal dts As vcliente) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_vendedor")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@vendedor", dts.gVendedor)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@fcom", dts.gfcom)
            cmd.Parameters.AddWithValue("@fcom_fin", dts.gfcom_fin)
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
    Public Function buscar_cliente(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_cliente")
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
    Public Function buscar_cliente_ruc(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_cliente_ruc")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
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
    Public Function mostrar_cliente(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_cliente")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gVendedor)
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
    Public Function buscar_cliente20(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_cliente20")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cia", dts.gruc)
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
    Public Function buscar_cliente_vianny(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_cliente_vianny")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ficha", dts.gruc)
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
    Public Function buscar_direccion(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_direccion")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ubigeo", dts.gubigeo)
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
    Public Function buscar_departamento() As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_departamento")
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
    Public Function buscar_PROVINCIA(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_PROVINCIA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@DEPARTAMENTO", dts.gdepartamento)
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
    Public Function BUSCAR_DISTRITO(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_DISTRITO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@PROVINCIA", dts.gprovincia)
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

    Public Function CLIENTE_FRANCISCO(ByVal dts As vcliente) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CLIENTE_FRANCISCO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@RUC", dts.gruc)
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
    Public Function BUSCAR_VENDEDOR_CLIENTE(ByVal dts As vcliente)
        Try
            conectare()
            cmd = New SqlCommand("BUCAR_VENDEDOR_CLIENTE")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CLIENTE", dts.gruc)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("Vendedor").ToString()
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
    Public Function BUSCAR_cliente_gfff(ByVal dts As vcliente)
        Try
            conectare()
            cmd = New SqlCommand("buscar_cliente_vianny")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ficha", dts.gruc)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("nomb_10").ToString()
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
    Public Function eliminar_CLIENTE(ByVal dts As vcliente) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ELIMINAR_CLIENTE")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CODIGO", dts.gcodigo)

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
