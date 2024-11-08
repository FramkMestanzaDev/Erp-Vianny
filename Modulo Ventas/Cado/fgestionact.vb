Imports System.Data.SqlClient
Public Class fgestionact
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_gestoosn_actividades(ByVal dts As vgestionact) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_gestion_actividades")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@id_ga", dts.gid_ga)
            cmd.Parameters.AddWithValue("@fecha_g", dts.gfecha_g)
            cmd.Parameters.AddWithValue("@vededor_co", dts.gvededor_co)
            cmd.Parameters.AddWithValue("@vendedor_nom", dts.gvendedor_nom)
            cmd.Parameters.AddWithValue("@cliente_ruc", dts.gcliente_ruc)
            cmd.Parameters.AddWithValue("@cliente_rs", dts.gcliente_rs)
            cmd.Parameters.AddWithValue("@prox_visita", dts.gprox_visita)
            cmd.Parameters.AddWithValue("@contacto", dts.gcontacto)
            cmd.Parameters.AddWithValue("@estado", dts.gestado)
            cmd.Parameters.AddWithValue("@descripcion", dts.gdescripcion)
            cmd.Parameters.AddWithValue("@ubicacion", dts.gubicacion)
            cmd.Parameters.AddWithValue("@CORREO", dts.gCORREO)
            cmd.Parameters.AddWithValue("@CELULAR", dts.gCELULAR)
            cmd.Parameters.AddWithValue("@MONEDA", dts.gMONEDA)
            cmd.Parameters.AddWithValue("@MONTO", dts.gMONTO)
            cmd.Parameters.AddWithValue("@CONTACTARON", dts.gCONTACTARON)
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
    Public Function correlativo_GESTION()
        Try
            conectare()
            cmd = New SqlCommand("correlativo_GESTIO_ACTIVI")
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
    Public Function mostrar_gestion_actividades(ByVal dts As vgestionact) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_gestion_actividades")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvededor_co)
            cmd.Parameters.AddWithValue("@fecha_ini", dts.gfini)
            cmd.Parameters.AddWithValue("@fecha_fin", dts.gffin)
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
    Public Function mostrar_gestion_actividades2(ByVal dts As vgestionact) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_gestion_actividades2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@vendedor", dts.gvededor_co)
            cmd.Parameters.AddWithValue("@fecha_ini", dts.gfini)
            cmd.Parameters.AddWithValue("@fecha_fin", dts.gffin)
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
