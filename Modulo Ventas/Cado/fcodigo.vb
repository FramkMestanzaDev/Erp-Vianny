Imports System.Data.SqlClient
Public Class fcodigo
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_codigo(ByVal dts As vcodigo) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_codigo_producto")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@familia", dts.gfamilia)
            cmd.Parameters.AddWithValue("@subfamilia", dts.gsubfamilia)
            cmd.Parameters.AddWithValue("@codcolor", dts.gcodcolor)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@ancho", dts.gancho)
            cmd.Parameters.AddWithValue("@densidad", dts.gdensidad)
            cmd.Parameters.AddWithValue("@detalle", dts.gdetalle)
            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@nombre_comercial", dts.gnombre_comercial)
            cmd.Parameters.AddWithValue("@um", dts.gum)
            cmd.Parameters.AddWithValue("@TipCtrl_17", dts.gTipCtrl_17)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@tipo", dts.gtipo)
            cmd.Parameters.AddWithValue("@barra", dts.gbarra)
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
    Public Function insertar_codigo_planemiento(ByVal dts As vcodigo) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_codigo_planemiento")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@linea2", dts.glinea2)
            cmd.Parameters.AddWithValue("@linea", dts.glinea)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
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
    Public Function buscar_informacion(ByVal dts As vlogica) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.CaeSoft_GetAllMaximoProducto")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gCCia)
            cmd.Parameters.AddWithValue("@Familia", dts.gFamilia)
            cmd.Parameters.AddWithValue("@SubFamilia", dts.gSubFamilia)
            cmd.Parameters.AddWithValue("@Origen", dts.gOrigen)
            cmd.Parameters.AddWithValue("@Color", dts.gColor)
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
    Public Function mostrar_correlativo(ByVal dts As vlogica)
        Try
            conectare()
            cmd = New SqlCommand("custom_vianny.dbo.CaeSoft_GetAllMaximoProducto")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@CCia", dts.gCCia)
            cmd.Parameters.AddWithValue("@Familia", dts.gFamilia)
            cmd.Parameters.AddWithValue("@SubFamilia", dts.gSubFamilia)
            cmd.Parameters.AddWithValue("@Origen", dts.gOrigen)
            cmd.Parameters.AddWithValue("@Color", dts.gColor)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("Maximo").ToString()
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
