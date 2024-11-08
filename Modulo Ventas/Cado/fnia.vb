Imports System.Data.SqlClient
Public Class fnia
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_cabecera_nia(ByVal dts As vinsertar_nia) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertarcabecera_nia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@glosa", dts.gglosa)
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@guia", dts.gguia)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@origen", dts.gorigen)
            cmd.Parameters.AddWithValue("@int", dts.gint)
            cmd.Parameters.AddWithValue("@empresa", dts.gempresa)
            cmd.Parameters.AddWithValue("@tidoc", dts.gtidoc)
            cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@fase", dts.gfase)
            If cmd.ExecuteNonQuery Then
                Return True
                MsgBox(dts.gfecha)
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
    Public Function mostrar_detalle_nia_pt(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_detalle_nia_pt")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcomprobante)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen1)
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
    Public Function eliminar_nia(ByVal dts As vnia) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_nia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcomprobante)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@compro", dts.gncom)
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
    Public Function insertar_detalle_nia(ByVal dts As insertardetallenia) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertardetalle_nia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
            cmd.Parameters.AddWithValue("@item", dts.gitem)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@cantidad789", dts.gcantidad1)
            cmd.Parameters.AddWithValue("@und", dts.gund)
            cmd.Parameters.AddWithValue("@op", dts.gop)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@umpresentacion", dts.gumpresentacion)
            cmd.Parameters.AddWithValue("@otejeduria", dts.gotejeduria)
            cmd.Parameters.AddWithValue("@oc", dts.goc)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@cenvio", dts.gcenvio)
            cmd.Parameters.AddWithValue("@clasif", dts.gclasif)
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
    Public Function buscar_guia(ByVal dts As vnia)
        Try
            conectare()
            cmd = New SqlCommand("guia_ingresada")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@guia", dts.gguia)
            cmd.Parameters.AddWithValue("@ano", dts.gano)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.HasRows = True Then
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
    Public Function buscar_nia(ByVal dts As vnia)
        Try
            conectare()
            cmd = New SqlCommand("correla_nia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
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
    Public Function mostrar_detalle_nia(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_detalle_nia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcomprobante)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen1)
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
    Public Function PARTE_INGRESO(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("PARTE_INGRESO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@PARTIDA", dts.gpartida)
            cmd.Parameters.AddWithValue("@LINEA", dts.glinea)
            cmd.Parameters.AddWithValue("@MES", dts.gmes)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@CCOM", dts.gccom)
            cmd.Parameters.AddWithValue("@po", dts.gop)
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
    Public Function TABLA_PRODUCTOS(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("TABLA_PRODUCTOS")
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
    Public Function mostrar_cabecera_nia(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_cabecera_nia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcomprobante)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen1)
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
    Public Function mostrar_correlativo_nia1(ByVal dts As vnia)
        Try
            conectare()
            cmd = New SqlCommand("buscar_nia")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcomprobante)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen1)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
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
    Public Function anular_nia(ByVal dts As vnia) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("anular_nia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@comprobante", dts.gcomprobante)
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
End Class
