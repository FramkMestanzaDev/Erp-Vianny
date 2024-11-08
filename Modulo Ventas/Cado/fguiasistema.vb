Imports System.Data.SqlClient
Public Class fguiasistema
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_cabecera_guia_sistema(ByVal dts As vguiacabecera) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertarcabecera_guia_sistema")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)
            cmd.Parameters.AddWithValue("@ruc", dts.gruc)
            cmd.Parameters.AddWithValue("@rsocial", dts.grsocial)
            cmd.Parameters.AddWithValue("@fcom", dts.gfcom)
            cmd.Parameters.AddWithValue("@fcom1", dts.gfcom1)
            cmd.Parameters.AddWithValue("@direccion", dts.gdireccion)
            cmd.Parameters.AddWithValue("@placa", dts.gplaca)
            cmd.Parameters.AddWithValue("@direccion_partida", dts.gdireccion_partida)
            cmd.Parameters.AddWithValue("@NOTASALIDA", dts.gNOTASALIDA)
            cmd.Parameters.AddWithValue("@chofer", dts.gchofer)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            cmd.Parameters.AddWithValue("@correlativo", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@licencia", dts.glicencia)
            cmd.Parameters.AddWithValue("@ruc_transportista", dts.gruc_transpo)
            cmd.Parameters.AddWithValue("@direc_transport", dts.gdirec_transport)
            cmd.Parameters.AddWithValue("@tip_documento", dts.gtip_documento)
            cmd.Parameters.AddWithValue("@ruc3", dts.gruc3)
            cmd.Parameters.AddWithValue("@motivo", dts.gmotivo)
            cmd.Parameters.AddWithValue("@destino", dts.gdestino)
            cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
            cmd.Parameters.AddWithValue("@fase", dts.gfase)
            cmd.Parameters.AddWithValue("@trpo", dts.gtrpo)
            cmd.Parameters.AddWithValue("@glosa", dts.gglosa)
            cmd.Parameters.AddWithValue("@ordens_3", dts.gordens_3)
            cmd.Parameters.AddWithValue("@subfase_3", dts.gsubfase_3)
            cmd.Parameters.AddWithValue("@situacion", dts.gsituacion)
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
    Public Function insertar_detalle_guia_sistema(ByVal dts As vguiadetalle) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertardetalle_guia_sistema")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@sfactu", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nfactu", dts.gnfactu)
            cmd.Parameters.AddWithValue("@items", dts.gitems)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@ordens", dts.gordens)
            cmd.Parameters.AddWithValue("@rollo", dts.grollo)
            cmd.Parameters.AddWithValue("@partida", dts.gparti)
            cmd.Parameters.AddWithValue("@unidad_medida", dts.gunidad_medida)
            cmd.Parameters.AddWithValue("@unidad_medida2", dts.gunidad_medida2)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@canencio", dts.gcanencio)
            cmd.Parameters.AddWithValue("@ordtejeduria", dts.gordtejeduria)
            cmd.Parameters.AddWithValue("@obser1_3a", dts.gobser1_3a)
            cmd.Parameters.AddWithValue("@obser2_3a", dts.gobser2_3a)
            cmd.Parameters.AddWithValue("@obser3_3a", dts.gobser3_3a)
            cmd.Parameters.AddWithValue("@obser4_3a", dts.gobser4_3a)
            cmd.Parameters.AddWithValue("@obser5_3a", dts.gobser5_3a)
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
    Public Function mostrar_guia_correlativo(ByVal dts As vguiacabecera)
        Try
            conectare()
            cmd = New SqlCommand("correla_guia")
            cmd.Parameters.AddWithValue("@ccia ", dts.gccia)
            cmd.Parameters.AddWithValue("@serie ", dts.gserie)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("nfactu_3").ToString()
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
    Public Function mostrar_almacen_guia(ByVal dts As vguiacabecera)
        Try
            conectare()
            cmd = New SqlCommand("buscar_almacen guia")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("almreg_3").ToString()
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
    Public Function buscar_ano_guia(ByVal dts As vguiacabecera)
        Try
            conectare()
            cmd = New SqlCommand("buscar_ano_guia")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ano").ToString()
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
    Public Function mostrar_almacen_guia_prod(ByVal dts As vguiacabecera)
        Try
            conectare()
            cmd = New SqlCommand("buscar_almacen_prod")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fasereg_3").ToString()
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
    Public Function mostrar_correlativo_nsa(ByVal dts As vguiacabecera)
        Try
            conectare()
            cmd = New SqlCommand("buscar_correlativo_nsa")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ncom_3").ToString()
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
    Public Function mostrar_cabecera_guia(ByVal dts As vguiacabecera) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_cabecera_guia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@correlativo", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
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
    Public Function mostrar_cabecera_guia_prod(ByVal dts As vguiacabecera) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_cabecera_guia_prod")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@correlativo", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
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
    Public Function guia_mostrar(ByVal dts As vguiacabecera) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("guia_mostrar")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            cmd.Parameters.AddWithValue("@fechaini", dts.gfini)
            cmd.Parameters.AddWithValue("@fechafin", dts.gffin)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
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
    Public Function stock_guia(ByVal dts As vguiacabecera) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteStockFisico_38")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gccia)
            cmd.Parameters.AddWithValue("@Almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@CodArtIni", dts.gCodArtIni)
            If Trim(dts.gFiltroDescrip) = "1" Then
                cmd.Parameters.AddWithValue("@FiltroDescrip", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@FiltroDescrip", dts.gFiltroDescrip)
            End If

            cmd.Parameters.AddWithValue("@Modal", dts.gModal)
            cmd.Parameters.AddWithValue("@ano", dts.gano)
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

    Public Function mostrar_detalle_guia(ByVal dts As vguiacabecera) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_detalle_guia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
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
    Public Function mostrar_detalle_guia_hilo(ByVal dts As vguiacabecera) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("mostrar_detalle_guia_hilo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
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
    Public Function mostrar_correlativo_guia1(ByVal dts As vguiacabecera)
        Try
            conectare()
            cmd = New SqlCommand("buscar_guia1")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@CORRELATIVO", dts.gcorrelativo)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
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
    Public Function eliminar_guia(ByVal dts As vguiacabecera) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_guia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@sgui", dts.gsfactu)
            cmd.Parameters.AddWithValue("@nguia", dts.gnfactu)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
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
    Public Function anular_guia(ByVal dts As vguiacabecera) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("anular_guia")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@serie", dts.gsfactu)
            cmd.Parameters.AddWithValue("@correlativo", dts.gnfactu)
            cmd.Parameters.AddWithValue("@ano", dts.gano)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@partida", dts.gpartida)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
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
    Public Function mostrar_sinonimo(ByVal dts As vguiacabecera)
        Try
            conectare()
            cmd = New SqlCommand("buscar_correlativo_nsa")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@mes", dts.gmes)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ncom_3").ToString()
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
