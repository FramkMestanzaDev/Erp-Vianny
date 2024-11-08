Imports System.Data.SqlClient
Public Class faccion
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function insertar_acciones(ByVal dts As vacciones) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insert_Registro_acciones")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@Accion_realiza", dts.gAccion_realiza)
            cmd.Parameters.AddWithValue("@Nombre_fomulario", dts.gNombre_fomulario)
            cmd.Parameters.AddWithValue("@Codigo_iden", dts.gCodigo_iden)
            cmd.Parameters.AddWithValue("@usuario", dts.gusuario)
            cmd.Parameters.AddWithValue("@fyh", dts.gfyh)
            cmd.Parameters.AddWithValue("@maquina", dts.gmaquina)


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
