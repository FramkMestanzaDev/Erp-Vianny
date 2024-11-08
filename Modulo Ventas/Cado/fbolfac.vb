Imports System.Data.SqlClient
Public Class fbolfac
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function mostrar_correlativo(ByVal dts As vdocfac)
        Try
            conectare()
            cmd = New SqlCommand("buscar_fac_bol")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fac").ToString()
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
    Public Function mostrar_correlativo_2(ByVal dts As vdocfac)
        Try
            conectare()
            cmd = New SqlCommand("buscar_fac_bol_2")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@almacen", dts.galmacen)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fac").ToString()
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

    Public Function mostrar_correlativo_3(ByVal dts As vdocfac)
        Try
            conectare()
            cmd = New SqlCommand("buscar_fac_bol_3")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@doc", dts.gdoc)
            cmd.Parameters.AddWithValue("@serie", dts.gserie)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("fac").ToString()
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
