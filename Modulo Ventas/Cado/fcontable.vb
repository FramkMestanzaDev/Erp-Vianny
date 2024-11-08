Imports System.Data.SqlClient
Public Class fcontable
    Inherits conexion
    Dim cmd As New SqlCommand

    Public Function insertar_cabecera_contable(ByVal dts As vcontable) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("insertacabeceraregitro_contab")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia_3", dts.gccia_3)
            cmd.Parameters.AddWithValue("@cperiodo_3", dts.gcperiodo_3)
            cmd.Parameters.AddWithValue("@ccom_3", dts.gccom_3)
            cmd.Parameters.AddWithValue("@ncom_3", dts.gncom_3)
            cmd.Parameters.AddWithValue("@tcom_3", dts.gtcom_3)
            cmd.Parameters.AddWithValue("@mone_3", dts.gmone_3)
            cmd.Parameters.AddWithValue("@fcom_3", dts.gfcom_3)
            cmd.Parameters.AddWithValue("@tcam_3", dts.gtcam_3)
            cmd.Parameters.AddWithValue("@glosa_3", dts.gglosa_3)
            cmd.Parameters.AddWithValue("@glosb_3", dts.gglosb_3)
            cmd.Parameters.AddWithValue("@totd1_3", dts.gtotd1_3)
            cmd.Parameters.AddWithValue("@totd2_3", dts.gtotd2_3)
            cmd.Parameters.AddWithValue("@toth1_3", dts.gtoth1_3)
            cmd.Parameters.AddWithValue("@toth2_3", dts.gtoth2_3)
            cmd.Parameters.AddWithValue("@cuenb_3", dts.gcuenb_3)
            cmd.Parameters.AddWithValue("@girado_3", dts.ggirado_3)
            cmd.Parameters.AddWithValue("@nombcp_3", dts.gnombcp_3)
            cmd.Parameters.AddWithValue("@fcoma_3", dts.gfcoma_3)
            cmd.Parameters.AddWithValue("@nroa_3", dts.gnroa_3)
            cmd.Parameters.AddWithValue("@tmov_3", dts.gtmov_3)
            cmd.Parameters.AddWithValue("@difcam_3", dts.gdifcam_3)
            cmd.Parameters.AddWithValue("@tasien_3", dts.gtasien_3)
            cmd.Parameters.AddWithValue("@flag_3", dts.gflag_3)
            cmd.Parameters.AddWithValue("@fcome_3", dts.gfcome_3)
            cmd.Parameters.AddWithValue("@usuario_3", dts.gusuario_3)
            cmd.Parameters.AddWithValue("@fupdate_3", dts.gfupdate_3)
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
    Public Function BUSCAR_SUBDIARIO(ByVal dts As vcontable)
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_SUBDIARIO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn

            cmd.Parameters.AddWithValue("@ccia", dts.gccia_3)
            cmd.Parameters.AddWithValue("@ALMACEN", dts.gALMACEN)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("SUB").ToString()
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
    Public Function BUSCAR_RUBRO(ByVal dts As vcontable)
        Try
            conectare()
            cmd = New SqlCommand("BUSCAR_RUBRO")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia_3)
            cmd.Parameters.AddWithValue("@RUBRO", dts.gRUBRO)

            Dim dr As SqlDataReader

            dr = cmd.ExecuteReader
            If dr.Read = True Then
                Return dr("CUENTA").ToString()
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
    Public Function insertar_detallecontable(ByVal dts As vdeta_contable) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("inertar_detalle_regcuenta")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn


            cmd.Parameters.AddWithValue("@ccia_3a", dts.gccia_3a)
            cmd.Parameters.AddWithValue("@cperiodo_3a", dts.gcperiodo_3a)
            cmd.Parameters.AddWithValue("@ccom_3a", dts.gccom_3a)
            cmd.Parameters.AddWithValue("@ncom_3a", dts.gncom_3a)
            cmd.Parameters.AddWithValue("@cuen_3a", dts.gcuen_3a)
            cmd.Parameters.AddWithValue("@cuend_3a", dts.gcuend_3a)
            cmd.Parameters.AddWithValue("@cuenc_3a", dts.gcuenc_3a)
            cmd.Parameters.AddWithValue("@codh_3a", dts.gcodh_3a)
            cmd.Parameters.AddWithValue("@fich_3a", dts.gfich_3a)
            cmd.Parameters.AddWithValue("@cdoc_3a", dts.gcdoc_3a)
            cmd.Parameters.AddWithValue("@ndoc_3a", dts.gndoc_3a)
            cmd.Parameters.AddWithValue("@cref_3a", dts.gcref_3a)
            cmd.Parameters.AddWithValue("@nref_3a", dts.gnref_3a)
            cmd.Parameters.AddWithValue("@linea_3a", dts.glinea_3a)
            cmd.Parameters.AddWithValue("@pptof_3a", dts.gpptof_3a)
            cmd.Parameters.AddWithValue("@nomb_3a", dts.gnomb_3a)
            cmd.Parameters.AddWithValue("@glosa_3a", dts.gglosa_3a)
            cmd.Parameters.AddWithValue("@imp1_3a", dts.gimp1_3a)
            cmd.Parameters.AddWithValue("@imp2_3a", dts.gimp2_3a)
            cmd.Parameters.AddWithValue("@fref_3a", dts.gfref_3a)
            cmd.Parameters.AddWithValue("@flagd_3a", dts.gflagd_3a)
            cmd.Parameters.AddWithValue("@proy_3a", dts.gproy_3a)
            cmd.Parameters.AddWithValue("@princi_3a", dts.gprinci_3a)
            cmd.Parameters.AddWithValue("@fuen_3a", dts.gfuen_3a)
            cmd.Parameters.AddWithValue("@parti_3a", dts.gparti_3a)
            cmd.Parameters.AddWithValue("@tefe_3a", dts.gtefe_3a)
            cmd.Parameters.AddWithValue("@sitlet_3a", dts.gsitlet_3a)
            cmd.Parameters.AddWithValue("@banco_3a", dts.gbanco_3a)
            cmd.Parameters.AddWithValue("@nrounico_3a", dts.gnrounico_3a)
            cmd.Parameters.AddWithValue("@endoso_3a", dts.gendoso_3a)
            cmd.Parameters.AddWithValue("@aval_3a", dts.gaval_3a)
            cmd.Parameters.AddWithValue("@flag_3a", dts.gflag_3a)
            cmd.Parameters.AddWithValue("@ocompra_3a", dts.gocompra_3a)
            cmd.Parameters.AddWithValue("@ordens_3a", dts.gordens_3a)
            cmd.Parameters.AddWithValue("@tcam_3a", dts.gtcam_3a)
            cmd.Parameters.AddWithValue("@femision_3a", dts.gfemision_3a)

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
    Public Function eliminar_registro_contable(ByVal dts As vcontable) As Boolean
        Try
            conectare()
            cmd = New SqlCommand("eliminar_registro_contable")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = cnn
            cmd.Parameters.AddWithValue("@ccia", dts.gccia_3)
            cmd.Parameters.AddWithValue("@cperiodo", dts.gcperiodo_3)
            cmd.Parameters.AddWithValue("@ccom", dts.gccom_3)
            cmd.Parameters.AddWithValue("@ncom", dts.gncom_3)

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
