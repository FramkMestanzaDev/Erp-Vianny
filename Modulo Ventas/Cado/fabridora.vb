Imports System.Data.SqlClient
Public Class fabridora
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function correlativo_abridora()
        Try
            conectare()
            cmd = New SqlCommand("correlativo_abridora")
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
    Public Function insertar_abridora_cabecera(ByVal dts As vabridora) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_abridora_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@id_abridora", dts.gid)
            cmd.Parameters.AddWithValue("@juego_urdido", dts.gjuego_urdido)
            cmd.Parameters.AddWithValue("@juego_abridora", dts.gjuego_abridora)
            cmd.Parameters.AddWithValue("@juego_tenido", dts.gjuego_tenido)
            cmd.Parameters.AddWithValue("@juego_conera", dts.gjuego_conera)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo)
            cmd.Parameters.AddWithValue("@titulo", dts.gtitulo)
            cmd.Parameters.AddWithValue("@hilo", dts.ghilo)
            cmd.Parameters.AddWithValue("@longitud", dts.glongitud)
            cmd.Parameters.AddWithValue("@lote", dts.glote)
            cmd.Parameters.AddWithValue("@cuerdas", dts.gcuerda)

            cmd.Parameters.AddWithValue("@flag", dts.gflag)
            cmd.Parameters.AddWithValue("@programa", dts.gprograma)
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
    Public Function insertar_abridora_detalle(ByVal dts As vabridora) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_abridora_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@id_detalle_abridora", dts.gid)
            cmd.Parameters.AddWithValue("@juego_abridora", dts.gjuego_abridorad)
            cmd.Parameters.AddWithValue("@cuerda", dts.gcuerdad)
            cmd.Parameters.AddWithValue("@tacho", dts.gtachod)
            cmd.Parameters.AddWithValue("@maq", dts.gmaquinad)
            cmd.Parameters.AddWithValue("@plegador", dts.gplegadord)
            cmd.Parameters.AddWithValue("@fecha", dts.gfechad)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigo)
            cmd.Parameters.AddWithValue("@operario", dts.gtrabajador)
            cmd.Parameters.AddWithValue("@titulo", dts.gtitulod)
            cmd.Parameters.AddWithValue("@hilo", dts.ghilod)
            cmd.Parameters.AddWithValue("@lote", dts.gloted)
            cmd.Parameters.AddWithValue("@despacho", dts.gdespachod)
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
    Public Function buscar_abridora(ByVal dts As vabridora) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("buscar_abridora")
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
    Public Function eliminar_abridora(ByVal dts As vabridora) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("elimiar_abridora")
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
