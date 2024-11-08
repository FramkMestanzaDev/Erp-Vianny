Imports System.Data.SqlClient
Public Class fusuario
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_usuario(ByVal dts As vusuario) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_usuario")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@id_usuario", dts.gid_usuario)
            cmd.Parameters.AddWithValue("@nombre", dts.gnombre)
            cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
            cmd.Parameters.AddWithValue("@clave", dts.gclave)
            cmd.Parameters.AddWithValue("@grupo", dts.ggrupo)
            cmd.Parameters.AddWithValue("@correo", dts.gcorreo)

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

    Public Function validar_login(ByVal dts As vusuario)
        Try
            conectare()
            cmd = New SqlCommand("buscar_login")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
            cmd.Parameters.AddWithValue("@clave", dts.gclave)
            cmd.Parameters.AddWithValue("@grupo", dts.ggrupo)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("CANT").ToString()
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
    Public Function BUSCAR_GRUPO(ByVal dts As vusuario)
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_GRUPO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@USUARIO", dts.ggrupo)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("grupo").ToString()
            Else
                MsgBox("EL USUARIO NO EXISTE")
                Return "1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            desconectar()

        End Try
    End Function
End Class
