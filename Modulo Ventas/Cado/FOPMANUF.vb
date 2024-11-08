Imports System.Data.SqlClient

Public Class FOPMANUF
    Inherits conexion
    Dim cmd As New SqlCommand
    Public Function CORRELTIVO_PRODUCTO(ByVal dts As vopmanufac)
        Try
            conectare()
            cmd = New SqlCommand("CaeSoft_GetAllMaximoProducto")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@Familia", dts.gfamilia)
            cmd.Parameters.AddWithValue("@SubFamilia", dts.gSubFamilia)
            cmd.Parameters.AddWithValue("@Origen", dts.gOrigen)
            cmd.Parameters.AddWithValue("@Color", dts.gcolor)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("Maximo").ToString()
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

    Public Function insertar_op(ByVal dts As vopmanufac) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_op_manufactura")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom_3)
            cmd.Parameters.AddWithValue("@fich_3", dts.gfich_3)
            cmd.Parameters.AddWithValue("@fcom_3", dts.gfcom_3)
            cmd.Parameters.AddWithValue("@mone_3", dts.gmone_3)
            cmd.Parameters.AddWithValue("@familia", dts.gfamilia)
            cmd.Parameters.AddWithValue("@linea_3", dts.glinea_3)
            cmd.Parameters.AddWithValue("@tallador", dts.gtallador)
            cmd.Parameters.AddWithValue("@FCome1_3", dts.gFCome1_3)
            cmd.Parameters.AddWithValue("@frequerida_3", dts.gfrequerida_3)
            cmd.Parameters.AddWithValue("@tcam_3", dts.gtcam_3)
            cmd.Parameters.AddWithValue("@program_3", dts.gprogram_3)
            cmd.Parameters.AddWithValue("@flag_3", dts.gflag_3)
            cmd.Parameters.AddWithValue("@descri_3", dts.gdescri_3)
            cmd.Parameters.AddWithValue("@alterno_3", dts.galterno_3)
            cmd.Parameters.AddWithValue("@zona", dts.gzona)
            cmd.Parameters.AddWithValue("@estilo_3", dts.gestilo_3)
            cmd.Parameters.AddWithValue("@pfob_3", dts.gpfob_3)
            cmd.Parameters.AddWithValue("@cantp_3", dts.gcantp_3)
            cmd.Parameters.AddWithValue("@cants_3", dts.gcants_3)
            cmd.Parameters.AddWithValue("@mdvent_3", dts.gmdvent_3)
            cmd.Parameters.AddWithValue("@tipo_mercado", dts.gtipo_mercado)
            cmd.Parameters.AddWithValue("@merchan_3", dts.gmerchan_3)
            cmd.Parameters.AddWithValue("@fpago", dts.gfpago)
            cmd.Parameters.AddWithValue("@broker_3", dts.gbroker_3)
            cmd.Parameters.AddWithValue("@viatra", dts.gviatra)
            cmd.Parameters.AddWithValue("@division", dts.gdivision)
            cmd.Parameters.AddWithValue("@estilo_empresa", dts.gestilo_empresa)
            cmd.Parameters.AddWithValue("@marcacli", dts.gmarcacli)
            cmd.Parameters.AddWithValue("@tipo_tejido", dts.gtipo_tejido)
            cmd.Parameters.AddWithValue("@observ_3", dts.gobserv_3)
            cmd.Parameters.AddWithValue("@cod_color", dts.gcod_color)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@tela", dts.gtela)
            cmd.Parameters.AddWithValue("@tipped_3", dts.gtipped_3)
            cmd.Parameters.AddWithValue("@mruta1_3", dts.gmruta1_3)
            cmd.Parameters.AddWithValue("@tela1_3", dts.gtela1_3)
            cmd.Parameters.AddWithValue("@lavadoten_3", dts.glavadoten_3)
            cmd.Parameters.AddWithValue("@otros1_3", dts.gotros1_3)
            cmd.Parameters.AddWithValue("@otros2_3", dts.gotros2_3)
            cmd.Parameters.AddWithValue("@Od", dts.gOd)
            cmd.Parameters.AddWithValue("@OdV", dts.gOdV)
            cmd.Parameters.AddWithValue("@mruta7_3", dts.gmruta7_3)
            cmd.Parameters.AddWithValue("@mruta8_3", dts.gmruta8_3)
            cmd.Parameters.AddWithValue("@scobs1_3", dts.gscobs1_3)
            cmd.Parameters.AddWithValue("@scobs2_3", dts.gscobs2_3)
            cmd.Parameters.AddWithValue("@scobs3_3", dts.gscobs3_3)
            cmd.Parameters.AddWithValue("@modelista_3", dts.gmodelista_3)
            cmd.Parameters.AddWithValue("@analista_3", dts.ganalista_3)
            cmd.Parameters.AddWithValue("@mruta5_3", dts.gmruta5_3)
            cmd.Parameters.AddWithValue("@mruta6_3", dts.gmruta6_3)
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

    Public Function insertar_cag170l(ByVal dts As vresto_op) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_cag170l")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom_3)
            cmd.Parameters.AddWithValue("@linea_3", dts.glinea_3)
            cmd.Parameters.AddWithValue("@corel", dts.gcorel)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function insertar_qag160d(ByVal dts As vresto_op) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_qag160d")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom", dts.gncom)
            cmd.Parameters.AddWithValue("@correl", dts.gcorrel)
            cmd.Parameters.AddWithValue("@cantidad", dts.gcantidad)
            cmd.Parameters.AddWithValue("@codcom", dts.gcodcom)
            cmd.Parameters.AddWithValue("@codtol", dts.gcodtol)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function insertar_qag160c(ByVal dts As vresto_op) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_qag160c")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom4", dts.gncom4)
            cmd.Parameters.AddWithValue("@correl", dts.gcorrelq)
            cmd.Parameters.AddWithValue("@talla", dts.gtalla)
            cmd.Parameters.AddWithValue("@codtal", dts.gcodtal)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
            cmd.Parameters.AddWithValue("@ratio", dts.gratio)
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
    Public Function insertar_qag160b(ByVal dts As vresto_op) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_qag160b")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ncom5", dts.gncom4)
            cmd.Parameters.AddWithValue("@cor", dts.gcorrelq)
            cmd.Parameters.AddWithValue("@color", dts.gcolor)
            cmd.Parameters.AddWithValue("@cants_3", dts.gcantidad)
            cmd.Parameters.AddWithValue("@codcom", dts.gcodcom)
            cmd.Parameters.AddWithValue("@colort", dts.gcolorrt)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
    Public Function insertar_cag1700(ByVal dts As insertar_codigo) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertar_codigo")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@linea_17", dts.glinea_17)
            cmd.Parameters.AddWithValue("@nomb_17", dts.gnomb_17)
            cmd.Parameters.AddWithValue("@unid_17", dts.gunid_17)
            cmd.Parameters.AddWithValue("@familia_17", dts.gfamilia_17)
            cmd.Parameters.AddWithValue("@subfam_17", dts.gsubfam_17)
            cmd.Parameters.AddWithValue("@tallaje_17", dts.gtallaje_17)
            cmd.Parameters.AddWithValue("@origen_17", dts.gorigen_17)
            cmd.Parameters.AddWithValue("@idcolor_17", dts.gidcolor_17)
            cmd.Parameters.AddWithValue("@articulo_17", dts.garticulo_17)
            cmd.Parameters.AddWithValue("@dmarca_17", dts.gdmarca_17)
            cmd.Parameters.AddWithValue("@codprod_17", dts.gcodprod_17)
            cmd.Parameters.AddWithValue("@ncolor_17", dts.gncolor_17)
            cmd.Parameters.AddWithValue("@ccia", dts.gccia)
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
