Imports System.Data.SqlClient
Public Class furdido
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function buscar_correla_urdido()
        Try
            conectare()
            cmd = New SqlCommand("buscar_correla_urdido")
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
    Public Function insertar_urdido_cabecera(ByVal dts As vurdido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_urdido_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@id_urdido", dts.gid)
            cmd.Parameters.AddWithValue("@juego_urdido", dts.gjuego_urdido)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo)
            cmd.Parameters.AddWithValue("@titulo", dts.gtitulo)
            cmd.Parameters.AddWithValue("@hilo", dts.ghilo)
            cmd.Parameters.AddWithValue("@longitud", dts.glongitud)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@ball", dts.gball)

            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@flag", dts.gflag)
            cmd.Parameters.AddWithValue("@programa", dts.gprograma)
            cmd.Parameters.AddWithValue("@maquina", dts.gmaquina)
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
    Public Function insertar_urdido_detalle(ByVal dts As vurdido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_urdido_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@id_detalle_urdido", dts.gid)
            cmd.Parameters.AddWithValue("@juego_urdido", dts.gjuego_urdidod)
            cmd.Parameters.AddWithValue("@items", dts.gitemsd)
            cmd.Parameters.AddWithValue("@ball", dts.gballd)
            cmd.Parameters.AddWithValue("@fecha", dts.gfechad)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigod)
            cmd.Parameters.AddWithValue("@operario", dts.goperariod)
            cmd.Parameters.AddWithValue("@turno1", dts.gturno1d)
            cmd.Parameters.AddWithValue("@turno2", dts.gturno2d)
            cmd.Parameters.AddWithValue("@turno3", dts.gturno3d)
            cmd.Parameters.AddWithValue("@hrotos", dts.ghrotosd)
            cmd.Parameters.AddWithValue("@incio", dts.ginciod)
            cmd.Parameters.AddWithValue("@fin", dts.gfind)
            cmd.Parameters.AddWithValue("@bolsas", dts.gbolsasd)
            cmd.Parameters.AddWithValue("@conos", dts.gconosd)
            cmd.Parameters.AddWithValue("@ayudante", dts.gayudanted)
            cmd.Parameters.AddWithValue("@flag", dts.gflagd)
            cmd.Parameters.AddWithValue("@programa", dts.gprogramad)


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
    Public Function buscar_urdido(ByVal dts As vurdido) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_urdido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@id", dts.gid)

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
    Public Function eliminar_urdido(ByVal dts As vurdido) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_urdido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@id", dts.gid)

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
