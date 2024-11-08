Imports System.Data.SqlClient
Public Class fprogramatenido
    Inherits conexion
    Dim cmd As New SqlCommand

    Public Function insertar_programa(ByVal dts As vprohgrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_programa")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@programa", dts.gprograma)
            cmd.Parameters.AddWithValue("@cargadas", dts.gcargadas)
            cmd.Parameters.AddWithValue("@fecha", dts.gfecha)
            cmd.Parameters.AddWithValue("@kg", dts.gkg)
            cmd.Parameters.AddWithValue("@flag", dts.gflag)


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
    Public Function insertar_programa_detalle(ByVal dts As vprohgrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_programa_detalle")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@abridora", dts.gabridorad)
            cmd.Parameters.AddWithValue("@urdido", dts.gurdidod)
            cmd.Parameters.AddWithValue("@conera", dts.gconerad)
            cmd.Parameters.AddWithValue("@tenido", dts.gtenidod)
            cmd.Parameters.AddWithValue("@programa", dts.gprogramad)
            cmd.Parameters.AddWithValue("@flag", dts.gflagd)
            cmd.Parameters.AddWithValue("@articulo", dts.garticulo)
            cmd.Parameters.AddWithValue("@titulo", dts.gtitulo)
            cmd.Parameters.AddWithValue("@hilo", dts.ghilo)

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
    Public Function eliminar_programa(ByVal dts As vprohgrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_programa_tenido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@programa", dts.gprograma)
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
    Public Function anular_program_tenido(ByVal dts As vprohgrama) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("anular_program_tenido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@programa", dts.gprograma)
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
    Public Function correlativo_programa()
        Try
            conectare()
            cmd = New SqlCommand("correlativo_programa_tenido")
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
    Public Function buscar_programa(ByVal dts As vprohgrama) As DataTable
        Try
            conectare()
            cmd = New SqlCommand("bucar_programa_tenido")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@programa", dts.gprograma)
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
