Imports System.Data.SqlClient
Public Class ftenidocuerda
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_tenido_cabecera(ByVal dts As vtenidocuerda) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_tenido_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("id_tenido", dts.gid)
            cmd.Parameters.AddWithValue("@juego_urdido", dts.gjuego_urdido)
            cmd.Parameters.AddWithValue("@juego_abridora", dts.gjuego_abridora)
            cmd.Parameters.AddWithValue("@juego_tenido", dts.gjuego_tenido)
            cmd.Parameters.AddWithValue("@juego_conera", dts.gjuego_conera)
            cmd.Parameters.AddWithValue("@articulo", dts.garticuloAs)
            cmd.Parameters.AddWithValue("@titulo", dts.gtitulo)
            cmd.Parameters.AddWithValue("@hilo", dts.ghilo)
            cmd.Parameters.AddWithValue("@longitud", dts.glongitud)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@trabajador", dts.gtrabajador)
            cmd.Parameters.AddWithValue("@flag", dts.gflag)
            cmd.Parameters.AddWithValue("@programa", dts.gprograma)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservacion)
            cmd.Parameters.AddWithValue("@velocidad", dts.gvelocidad)
            cmd.Parameters.AddWithValue("@ball", dts.gball)
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
    Public Function insertar_tenido_detalle(ByVal dts As vtenidocuerda) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_tenido_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@id_detalle_tenido", dts.gid)
            cmd.Parameters.AddWithValue("@juego_tenido", dts.gjuego_tenidod)
            cmd.Parameters.AddWithValue("@cuerda", dts.gcuerdad)
            cmd.Parameters.AddWithValue("@ball", dts.gballd)
            cmd.Parameters.AddWithValue("@tacho", dts.gtachod)
            cmd.Parameters.AddWithValue("@titulo", dts.gtitulod)
            cmd.Parameters.AddWithValue("@hilo", dts.ghilod)
            cmd.Parameters.AddWithValue("@lote", dts.gloted)
            cmd.Parameters.AddWithValue("@flag", dts.gflagd)
            cmd.Parameters.AddWithValue("@programa", dts.gprogramad)
            cmd.Parameters.AddWithValue("@observacion", dts.gobservaciond)
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
    Public Function buscar_tenidoooo(ByVal dts As vtenidocuerda) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_tenidoooo")
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
    Public Function eliminar_tenido_cuerda(ByVal dts As vtenidocuerda) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_tenido_cuerda")
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
    Public Function correlativo_teñido()
        Try
            conectare()
            cmd = New SqlCommand("correlativo_teñido")
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
End Class
