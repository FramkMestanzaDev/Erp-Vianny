Imports System.Data.SqlClient
Public Class fcontailidad
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar(ByVal dts As vcontablidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_registro_contabilidad")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn



            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@registro", dts.gregistro)
            cmd.Parameters.AddWithValue("@lacrado", dts.glacrado)
            cmd.Parameters.AddWithValue("@lacrado_tesoreria", dts.glacrado_tesoreria)
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
    Public Function eliminar(ByVal dts As vcontablidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_registro_contabilidad")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
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
    Public Function insertar_tesoreria(ByVal dts As vcontablidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_registro_tesoreria")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn



            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@registro", dts.gregistro)
            cmd.Parameters.AddWithValue("@lacrado", dts.glacrado)
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
    Public Function eliminar_TESORERIA(ByVal dts As vcontablidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_registro_TESORERIA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
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
    Public Function actualizar_tienda(ByVal dts As vcontablidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_rubro_tventa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cia", dts.gcia)
            cmd.Parameters.AddWithValue("@tienda", dts.gtienda)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@serieb", dts.gserieb)
            cmd.Parameters.AddWithValue("@serief", dts.gserief)
            cmd.Parameters.AddWithValue("@mes", dts.gMES)
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
    Public Function actualizar_tienda_graus(ByVal dts As vcontablidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_rubro_tventa2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@cia", dts.gcia)
            cmd.Parameters.AddWithValue("@tienda", dts.gtienda)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@serieb", dts.gserieb)
            cmd.Parameters.AddWithValue("@serief", dts.gserief)
            cmd.Parameters.AddWithValue("@mes", dts.gMES)
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
    Public Function actualizar_PAGO_VN(ByVal dts As vcontablidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("ACTUALIZAR_PG_PARCIAL")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@PARCIAL", dts.gPARCIAL)
            cmd.Parameters.AddWithValue("@OBSERVACION", dts.gOBSERVACION)
            cmd.Parameters.AddWithValue("@ID", dts.gID)
            cmd.Parameters.AddWithValue("@COMPROBANTE", dts.gCOMPROBANTE)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
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
    Public Function REPORTE_HUARACHA(ByVal dts As vcontablidad) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("reporte_huaracha")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@mes", dts.gMES)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
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
    Public Function REPORTE_HUARACHA2(ByVal dts As vcontablidad) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("REPORTE_HUARACHA_2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@mes", dts.gMES)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)

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
    Public Function buscar_factura_huaracha(ByVal dts As vcontablidad)
        Try
            conectare()
            cmd = New SqlCommand("buscar_factura_huaracha")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@grm", dts.ggrm)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fac").ToString()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function

    Public Function buscar_reportehh(ByVal dts As vcontablidad)
        Try
            conectare()
            cmd = New SqlCommand("buscar_reportehh")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fac").ToString()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function buscar_factura_fecha(ByVal dts As vcontablidad)
        Try
            conectare()
            cmd = New SqlCommand("buscar_fecha_factura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fecha").ToString()
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
    Public Function buscar_numdet(ByVal dts As vcontablidad)
        Try
            conectare()
            cmd = New SqlCommand("buscar_numdet")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@factura", dts.gfactura)
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@ccia", dts.gcia)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("numdet_3").ToString()
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
