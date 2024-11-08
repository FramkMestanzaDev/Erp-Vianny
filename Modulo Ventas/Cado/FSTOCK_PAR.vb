Imports System.Data.SqlClient
Public Class FSTOCK_PAR
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function actualizar_data_bancos(ByVal dts As VSTOCK_PAR)
        Try
            conectare()
            cmd = New SqlCommand("actualizar_bancos")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ocompra", dts.gocompra)
            cmd.Parameters.AddWithValue("@ccia", dts.gCCIA)
            cmd.Parameters.AddWithValue("@cuenta", dts.gcuenta)
            cmd.Parameters.AddWithValue("@registro", dts.gregistro)
            cmd.Parameters.AddWithValue("@periodo", dts.gperiodo)
            cmd.Parameters.AddWithValue("@dh", dts.gdh)
            cmd.Parameters.AddWithValue("@ccom", dts.gccom)
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
    Public Function actualiar_rubro(ByVal dts As VSTOCK_PAR)
        Try
            conectare()
            cmd = New SqlCommand("actualizar_rubro")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@ccia", dts.gCCIA)
            cmd.Parameters.AddWithValue("@rubro", dts.grubro)
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
    Public Function buscar_rubro(ByVal dts As VSTOCK_PAR)
        Try
            conectare()
            cmd = New SqlCommand("buscar_rubro12")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@ccia", dts.gCCIA)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("rubro_17").ToString()
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
    Public Function CaeSoft_ReporteStockFisico(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteStockFisico_37")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gCCIA)
            cmd.Parameters.AddWithValue("@Almacen", dts.gALMACEN)
            cmd.Parameters.AddWithValue("@Modal", dts.gMODAL)
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
    Public Function CaeSoft_ReporteStockFisico59(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteStockFisico_59")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gCCIA)
            cmd.Parameters.AddWithValue("@Almacen", dts.gALMACEN)
            cmd.Parameters.AddWithValue("@Modal", dts.gMODAL)
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
    Public Function CaeSoft_ReporteStockFisico_hilo(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteStockFisico_37_hilo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gCCIA)
            cmd.Parameters.AddWithValue("@Almacen", dts.gALMACEN)
            cmd.Parameters.AddWithValue("@Modal", dts.gMODAL)
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
    Public Function CaeSoft_ReporteStockFisico_3737(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_ReporteStockFisico_37_37")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gCCIA)
            cmd.Parameters.AddWithValue("@Almacen", dts.gALMACEN)
            cmd.Parameters.AddWithValue("@Modal", dts.gMODAL)
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
    Public Function CaeSoft_ReporteStockFisico_COMERCIAL(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("UNIR_PROCEDIMIENTO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@ccia", dts.gCCIA)
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
    Public Function CaeSoft_ReporteStockFisico_tercera(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("UNIR_PROCEDIMIENTO5")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@ccia", dts.gCCIA)
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

    Public Function CaeSoft_ReporteStockFisico_COMERCIAL_GRAUS(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("UNIR_PROCEDIMIENTO_GRAUS")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
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
    Public Function CaeSoft_ReporteStockFisico_COMERCIA2L(ByVal dts As VSTOCK_PAR) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("UNIR_PROCEDIMIENTO2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@periodo", dts.gano)
            cmd.Parameters.AddWithValue("@ccia", dts.gCCIA)
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
    Public Function mostrar_op(ByVal dts As VSTOCK_PAR)
        Try
            conectare()
            cmd = New SqlCommand("buscar_po")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@po", dts.gpo)
            cmd.Parameters.AddWithValue("@ccia", dts.gCCIA)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("ncom_3n").ToString()
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
End Class
