Imports System.Data.SqlClient
Public Class fpartida
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_calidad_partida(ByVal dts As vcalidadpartida) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_calidad_partida")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@color_des", dts.gcolor_des)
            cmd.Parameters.AddWithValue("@color_cod", dts.gcolor_cod)
            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@rollos", dts.grollos)
            cmd.Parameters.AddWithValue("@anchor", dts.ganchor)
            cmd.Parameters.AddWithValue("@densidadr", dts.gdensidadr)
            cmd.Parameters.AddWithValue("@lavadoa", dts.glavadoa)
            cmd.Parameters.AddWithValue("@lavadol", dts.glavadol)
            cmd.Parameters.AddWithValue("@lavador", dts.glavador)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@revisado", dts.grevisado)
            cmd.Parameters.AddWithValue("@adjuntado", dts.gadjuntado)
            cmd.Parameters.AddWithValue("@reproceso", dts.greproceso)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@elongacion", dts.gelongacion)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@merma", dts.gmerma)
            cmd.Parameters.AddWithValue("@maquina", dts.gmaquina)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@CENTRO_ORILLO", dts.gCENTRO_ORILLO)
            cmd.Parameters.AddWithValue("@DEGRADE", dts.gDEGRADE)
            cmd.Parameters.AddWithValue("@BARRADURA", dts.gBARRADURA)
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
    Public Function insertar_DETALLE_calidad_partida(ByVal dts As vdetallecalidad) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_detalle_calidad")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@DHCFM", dts.gDHCFM)
            cmd.Parameters.AddWithValue("@DHCPp", dts.gDHCPp)
            cmd.Parameters.AddWithValue("@DHCHG", dts.gDHCHG)
            cmd.Parameters.AddWithValue("@DHBF", dts.gDHBF)
            cmd.Parameters.AddWithValue("@DHTIH", dts.gDHTIH)
            cmd.Parameters.AddWithValue("@DTR", dts.gDTR)
            cmd.Parameters.AddWithValue("@DTLC", dts.gDTLC)
            cmd.Parameters.AddWithValue("@DTEL", dts.gDTEL)
            cmd.Parameters.AddWithValue("@DTRL", dts.gDTRL)
            cmd.Parameters.AddWithValue("@DTGA", dts.gDTGA)
            cmd.Parameters.AddWithValue("@DTLA", dts.gDTLA)
            cmd.Parameters.AddWithValue("@DTAG", dts.gDTAG)
            cmd.Parameters.AddWithValue("@DTCA", dts.gDTCA)
            cmd.Parameters.AddWithValue("@DTCT", dts.gDTCT)
            cmd.Parameters.AddWithValue("@DTAP", dts.gDTAP)
            cmd.Parameters.AddWithValue("@DTALR", dts.gDTALR)
            cmd.Parameters.AddWithValue("@DTATM", dts.gDTATM)
            cmd.Parameters.AddWithValue("@DTAML", dts.gDTAML)
            cmd.Parameters.AddWithValue("@DTAMV", dts.gDTAMV)
            cmd.Parameters.AddWithValue("@DTANMI", dts.gDTANMI)
            cmd.Parameters.AddWithValue("@DTAQT", dts.gDTAQT)
            cmd.Parameters.AddWithValue("@DTAPMS", dts.gDTAPMS)
            cmd.Parameters.AddWithValue("@DTARMB", dts.gDTARMB)
            cmd.Parameters.AddWithValue("@DTAMC", dts.gDTAMC)
            cmd.Parameters.AddWithValue("@DTAMG", dts.gDTAMG)
            cmd.Parameters.AddWithValue("@DTAMS", dts.gDTAMS)
            cmd.Parameters.AddWithValue("@DTAMCA", dts.gDTAMCA)
            cmd.Parameters.AddWithValue("@DTAM", dts.gDTAM)
            cmd.Parameters.AddWithValue("@DTAAG", dts.gDTAAG)
            cmd.Parameters.AddWithValue("@DTALMT", dts.gDTALMT)
            cmd.Parameters.AddWithValue("@DTALQ", dts.gDTALQ)
            cmd.Parameters.AddWithValue("@CC", dts.gCC)
            cmd.Parameters.AddWithValue("@CE", dts.gCE)
            cmd.Parameters.AddWithValue("@CDT", dts.gCDT)
            cmd.Parameters.AddWithValue("@DAR", dts.gDAR)
            cmd.Parameters.AddWithValue("@DO", dts.gDo1)
            cmd.Parameters.AddWithValue("@ROLLO", dts.gROLLO)
            cmd.Parameters.AddWithValue("@PARTIDA", dts.gPARTIDA)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@motas", dts.gmotas)
            cmd.Parameters.AddWithValue("@crama", dts.gcrama)
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
    Public Function buscar_partida_ingresada(ByVal dts As vpartida) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_CALIDAD_PARTIDA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@PARTIDA", dts.gpartida)
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
    Public Function buscar_partida_calidad(ByVal dts As vpartida) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_parti_calida")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
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
    Public Function buscar_rollos(ByVal dts As vpartida)
        Try
            conectare()
            cmd = New SqlCommand("buscar_rollos")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ROLLO").ToString()
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
    Public Function buscar_rollos_existe(ByVal dts As vpartida)
        Try
            conectare()
            cmd = New SqlCommand("buscar_rollo_existe")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@rollo", dts.grollo)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ROLLO").ToString()
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
    Public Function buscar_partida(ByVal dts As vpartida) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_parti")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
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
    Public Function buscar_persona(ByVal dts As vpartida)
        Try
            conectare()
            cmd = New SqlCommand("buscar_persona")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("noombre").ToString()
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
    Public Function buscar_partida2(ByVal dts As vpartida) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_PARTIDA2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
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
    Public Function buscar_codigo_packing(ByVal dts As vpartida) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_regis_guia")
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
    Public Function eliminar_calidad_partida(ByVal dts As vpartida) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_CALIDAD_PARTIDA")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@PARTIDA", dts.gpartida)

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
