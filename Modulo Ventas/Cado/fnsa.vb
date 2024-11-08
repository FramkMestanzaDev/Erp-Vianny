Imports System.Data.SqlClient
Public Class fnsa
    Inherits conexion
    Dim cmd As New SqlCommand

    Public Function anular_nSA(ByVal dts As vnsacabece) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("anular_nSa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@comprobante", dts.gncom)
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
    Public Function insertar_cabecera_nsa(ByVal dts As vnsacabece) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertarcabecera_nsa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@glosa", dts.gglosa)
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@guia", dts.gguia)
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
            cmd.Parameters.AddWithValue("@orige_sali", dts.gorige_sali)
            cmd.Parameters.AddWithValue("@tipointexte", dts.gtipointexte)
            cmd.Parameters.AddWithValue("@tdocento", dts.gtdocento)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@fase", dts.gfase)
            cmd.Parameters.AddWithValue("@adevol", dts.gadevol)
            cmd.Parameters.AddWithValue("@pedorig_3", dts.gpedorig_3)
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
    Public Function insertar_detalle_nsa(ByVal dts As vnsadetalle) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertardetalle_nsa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
            cmd.Parameters.AddWithValue("@item", dts.gitem)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@und", dts.gund)
            cmd.Parameters.AddWithValue("@op", dts.gop)
            cmd.Parameters.AddWithValue("@rollo", dts.grollo1)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@unidad_medidad", dts.gunidad_medidad)
            cmd.Parameters.AddWithValue("@ordtejeduria", dts.gordtejeduria)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@cantenvio", dts.gcantenvio)
            cmd.Parameters.AddWithValue("@clasif", dts.gclasif)
            cmd.Parameters.AddWithValue("@ocompra", dts.gOcompra)
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
    Public Function buscar_ot_po(ByVal dts As vnsadetalle) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_ot_po")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida ", dts.gpartida)
            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
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
    Public Function mostrar_cabecera(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_cabecera_nsa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
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
    Public Function mostrar_detalle(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("bucar_nsa_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
            cmd.Parameters.AddWithValue("@ncom ", dts.gncom)
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
    Public Function mostrar_correlativo_nsa1(ByVal dts As vnia)
        Try
            conectare()
            cmd = New SqlCommand("buscar_nsa")

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
    Public Function mostrar_cabecera_nsa(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_cabecera_nsa")
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
    Public Function mostrar_detalle_nsa(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_detalle_nsa")
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
    Public Function mostrar_detalle_nsa_pt(ByVal dts As vnia) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_detalle_nsa_pt")
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
    Public Function buscar_nsa(ByVal dts As vnia)
        Try
            conectare()
            cmd = New SqlCommand("correla_nsa")
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
End Class
