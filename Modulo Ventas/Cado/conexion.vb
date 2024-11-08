Imports System.Data.SqlClient
Public Class conexion
    Protected cnn As New SqlConnection
    Public id_usuario As Integer

    Public Function conectare()
        Try
            cnn = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=LAPTOP-CTM61GVR ;Initial Catalog=Comercial_Textil;integrated security = true")

            cnn.Open()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function desconectar()
        Try
            If cnn.State = ConnectionState.Open Then
                cnn.Close()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
End Class
