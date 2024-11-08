Imports System.Data.SqlClient
Public Class fingresopac
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function imprimir_packing(ByVal dts As vpackingtela) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("imprimir_packin_tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@calidad", dts.gnumero_rollos)
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
    Public Function imprimir_packing2(ByVal dts As vpackingtela) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("imprimir_packin_tela2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@almacen", dts.gnumero_rollos)
            cmd.Parameters.AddWithValue("@AL", dts.gAL)
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
    Public Function buscar_packing(ByVal dts As vpackingtela) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_packing_sina")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@calidad", dts.gnumero_rollos)
            cmd.Parameters.AddWithValue("@almacen", dts.galmac)
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
    Public Function eliminar_packing(ByVal dts As vpackingtela) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_packin_tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@calidad", dts.gnumero_rollos)
            cmd.Parameters.AddWithValue("@almacen", dts.galmac)
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
    Public Function insertar_packing_tela(ByVal dts As vpackingtela) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_pac_tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@des_articulo", dts.garticulo)
            cmd.Parameters.AddWithValue("@codigo_trab", dts.gcodigo_trab)
            cmd.Parameters.AddWithValue("@nombre_trab", dts.gnombre_trab)
            cmd.Parameters.AddWithValue("@numero_rollos", dts.gnumero_rollos)
            cmd.Parameters.AddWithValue("@unidad", dts.gunidad)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo3)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@total", dts.gtotal)
            cmd.Parameters.AddWithValue("@ROLLOS", dts.gROOLO)
            cmd.Parameters.AddWithValue("@seleccionado", dts.gseleccionado)
            cmd.Parameters.AddWithValue("@almac", dts.galmac)
            cmd.Parameters.AddWithValue("@orden", dts.gorden)
            cmd.Parameters.AddWithValue("@densidad", dts.gdesidad)
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
    Public Function insertar_packing_tela_detalle(ByVal dts As vpackingtela) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_det_pac_tela")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@nrollo", dts.gnrollo)
            cmd.Parameters.AddWithValue("@peso", dts.gpeso)
            cmd.Parameters.AddWithValue("@posicion", dts.gposicion)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo2)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@estado", dts.gestado1)
            cmd.Parameters.AddWithValue("@seleccionado", dts.gseleccionado)
            cmd.Parameters.AddWithValue("@idcabe", dts.gidcabe)
            cmd.Parameters.AddWithValue("@almac", dts.galmac)
            cmd.Parameters.AddWithValue("@peso_neto", dts.gpeso_neto)
            cmd.Parameters.AddWithValue("@ubic_art", dts.gubic_art)
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
    Public Function actualizar_packing(ByVal dts As vpackingtela) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualizar_packing")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@seleccionado", dts.gseleccionado)
            cmd.Parameters.AddWithValue("@guia", dts.gguia)
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
    Public Function actualizar_seleccionado(ByVal dts As vpackingtela) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualiza_estado_selecc")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@clasificacion", dts.gcalidad)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@es", dts.ges)
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
    Public Function actualizar_seleccionado_detalle(ByVal dts As vpackingtela) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("actualiza_estado_selecc_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@id", dts.gcodigo_det)
            cmd.Parameters.AddWithValue("@es", dts.ges)
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
    Public Function buscar_cant_packin(ByVal dts As vpackingtela)
        Try
            conectare()
            cmd = New SqlCommand("buscar_cant_packin")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("numero").ToString()
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
    Public Function ultimo_registro()
        Try
            conectare()
            cmd = New SqlCommand("ultimo_registro")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("id_cab").ToString()
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
    Public Function buscar_cant_packin_almacen(ByVal dts As vpackingtela)
        Try
            conectare()
            cmd = New SqlCommand("buscar_cant_packin_almacen")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("numero").ToString()
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
    Public Function buscar_rollo(ByVal dts As vpackingtela)
        Try
            conectare()
            cmd = New SqlCommand("buscar_rollo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@rollo", dts.gnumero_rollos)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("partida").ToString()
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
    Public Function peso_rollo(ByVal dts As vpackingtela)
        Try
            conectare()
            cmd = New SqlCommand("peso_rollo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@rollo", dts.gnumero_rollos)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("peso").ToString()
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
