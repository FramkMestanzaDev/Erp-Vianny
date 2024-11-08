Imports System.Data.SqlClient
Public Class fconera
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_abridora_cabecera(ByVal dts As vconera) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_abridora_cabecera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@id_conera", dts.gid_conera)
            cmd.Parameters.AddWithValue("@juego_urdido", dts.gjuego_urdido)
            cmd.Parameters.AddWithValue("@juego_abridora", dts.gjuego_abridora)
            cmd.Parameters.AddWithValue("@juego_tenido", dts.gjuego_tenido)
            cmd.Parameters.AddWithValue("@juego_conera", dts.gjuego_conera)
            cmd.Parameters.AddWithValue("@articulo", dts.gtitulo)
            cmd.Parameters.AddWithValue("@titulo", dts.ghilo)
            cmd.Parameters.AddWithValue("@hilo", dts.ghilo)
            cmd.Parameters.AddWithValue("@longitud", dts.glongitud)
            cmd.Parameters.AddWithValue("@conos", dts.gconos)
            cmd.Parameters.AddWithValue("@cuerda", dts.gcuerda)
            cmd.Parameters.AddWithValue("@hinicio", dts.ghinicio)
            cmd.Parameters.AddWithValue("@hfinal", dts.ghfinal)

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
    Public Function insertar_abridora_detalle(ByVal dts As vconera) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_detalle_conera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@id_conera", dts.gid_conerad)
            cmd.Parameters.AddWithValue("@codigo", dts.gcodigod)
            cmd.Parameters.AddWithValue("@operario", dts.goperariod)
            cmd.Parameters.AddWithValue("@fecha", dts.gfechad)
            cmd.Parameters.AddWithValue("@minicio", dts.gminiciod)
            cmd.Parameters.AddWithValue("@mfinal", dts.gmfinald)
            cmd.Parameters.AddWithValue("@conera", dts.gconerad)

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
    Public Function correlativo_conera()
        Try
            conectare()
            cmd = New SqlCommand("correlativo_conera")
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
    Public Function eliminar_conera(ByVal dts As vconera) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_conera")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@id", dts.gid_conera)

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
