Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class modulo_solicitud
    Public conx As SqlConnection
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim Rsr21, Rsr211, Rsr212, Rsr2127, Rsr215, Rsr210, Rsr2101, Rsr2105, Rsr2128, Rsr21289 As SqlDataReader
    Dim dg30, dg31 As New DataTable
    Sub confeccion_despachado()
        abrir()

        Dim cmd As New SqlDataAdapter("select id_requeminieto AS SOLICITUD, cliente AS CLIENTE,producto AS PRODUCTO,op AS OP,po AS PO,c_prog AS PROGRA,periodo AS PERIODO,corte AS CORTE,ISNULL( m.ncom_3,'') as REGISTRO, isnull(g.nomb_10,'') as PROVEEDOR
  from requerimiento_avios_cabecera c
  left join custom_vianny.dbo.mag030f  m on c.id_requeminieto = m.pedorig_3
  LEFT JOIN custom_vianny.dbo.cag1000  g on c.empresa = g.ccia and m.fich_3 = g.fich_10
  where estado IN (3)  and area=1 and periodo = YEAR(GETDATE())", conx)
        cmd.Fill(dg30)
        If dg30.Rows.Count <> 0 Then
            DataGridView1.DataSource = dg30

        Else
            DataGridView1.DataSource = ""
        End If

    End Sub
    Sub acabados_despachado()
        abrir()
        Dim cmd As New SqlDataAdapter("select id_requeminieto AS SOLICITUD, cliente AS CLIENTE,producto AS PRODUCTO,op AS OP,po AS PO,c_prog AS PROGRA,periodo AS PERIODO,corte AS CORTE,ISNULL( m.ncom_3,'') as REGISTRO, isnull(g.nomb_10,'') as PROVEEDOR
  from requerimiento_avios_cabecera c
  left join custom_vianny.dbo.mag030f  m on c.id_requeminieto = m.pedorig_3
  LEFT JOIN custom_vianny.dbo.cag1000  g on c.empresa = g.ccia and m.fich_3 = g.fich_10
  where estado IN (3)  and area=2 and periodo = YEAR(GETDATE())", conx)
        cmd.Fill(dg31)
        If dg31.Rows.Count <> 0 Then
            DataGridView1.DataSource = dg31

        Else
            DataGridView1.DataSource = ""
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).Visible = False
        Button6.Visible = True
        dg30.Clear()
        Label14.Text = "1"
        Button3.Enabled = False
        'DataGridView1.Rows.Clear()
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        If RadioButton4.Checked = True Then
            confeccion_despachado()
        Else
            If RadioButton3.Checked = True Then
                acabados_despachado()
            End If

        End If


    End Sub

    Private Sub buscar_op_confeccion()
        Try

            Dim ds As New DataSet
            ds.Tables.Add(dg1.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                DataGridView1.Columns(15).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(17).Visible = False
                DataGridView1.Columns(16).Width = 260
                DataGridView1.Columns(11).Width = 65
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar_op_acabados()
        Try

            Dim ds1 As New DataSet
            ds1.Tables.Add(dg2.Copy)
            Dim dv1 As New DataView(ds1.Tables(0))
            dv1.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"
            If dv1.Count <> 0 Then
                DataGridView1.DataSource = dv1
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                DataGridView1.Columns(15).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(17).Visible = False
                DataGridView1.Columns(16).Width = 260
                DataGridView1.Columns(11).Width = 65
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar_CONFECCION_DESPACHADO()
        Try
            Dim ds2 As New DataSet
            ds2.Tables.Add(dg30.Copy)
            Dim dv2 As New DataView(ds2.Tables(0))
            dv2.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"
            If dv2.Count <> 0 Then
                DataGridView1.DataSource = dv2
                DataGridView1.Columns(9).Visible = False
            Else
                DataGridView1.DataSource = Nothing
                DataGridView1.Columns(9).Visible = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar_DESPACHADO_DESPACHADO()
        Try
            Dim ds3 As New DataSet
            ds3.Tables.Add(dg31.Copy)
            Dim dv3 As New DataView(ds3.Tables(0))
            dv3.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"
            If dv3.Count <> 0 Then
                DataGridView1.DataSource = dv3

                DataGridView1.Columns(9).Visible = False

            Else
                DataGridView1.DataSource = Nothing

                DataGridView1.Columns(9).Visible = False

            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    'Private Sub buscar_PO()
    '    Try
    '        Dim ds As New DataSet
    '        ds.Tables.Add(dg1.Copy)
    '        Dim dv As New DataView(ds.Tables(0))
    '        dv.RowFilter = "PO" & " like '%" & TextBox2.Text & "%'"
    '        If dv.Count <> 0 Then
    '            DataGridView1.DataSource = dv
    '            DataGridView1.Columns(10).Visible = False
    '            DataGridView1.Columns(9).Visible = False
    '            DataGridView1.Columns(12).Visible = False
    '            DataGridView1.Columns(13).Visible = False
    '            DataGridView1.Columns(14).Visible = False
    '            DataGridView1.Columns(15).Visible = False
    '        Else
    '            DataGridView1.DataSource = Nothing
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message)

    '    End Try
    'End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Label14.Text = "1" Then
            If RadioButton4.Checked = True Then
                buscar_CONFECCION_DESPACHADO()
            Else
                If RadioButton3.Checked = True Then
                    buscar_DESPACHADO_DESPACHADO()
                End If
            End If
        Else
            If RadioButton1.Checked = True Then
                buscar_op_confeccion()
            Else
                If RadioButton2.Checked = True Then
                    buscar_op_acabados()


                End If
            End If


        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim a, conta, i As Integer
        conta = 0
        i = DataGridView1.Rows.Count
        For A = 0 To i - 1
            If Me.DataGridView1.Rows(A).Cells(0).Value = True Then
                conta = conta + 1
            End If
        Next

        If conta > 1 Then
            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
        Else
            Dim i1 As Integer
            Dim cco, fu As String
            If RadioButton1.Checked = True Then
                fu = "CONFECCION"
            Else
                fu = "ACABADOS"
            End If
            i1 = DataGridView1.Rows.Count
            For i = 0 To i1 - 1


                If DataGridView1.Rows(i).Cells(0).Value = True Then
                        abrir()
                    If Trim(DataGridView1.Rows(i).Cells(16).Value = "") Then
                        MsgBox("TODAVIA NO SE ASIGNADO UN TALLER A LA OP O PO")
                    Else
                        Dim respuesta As DialogResult
                        respuesta = MessageBox.Show("Generar Guia ? Si la respuesta es No Generara  una Nota de Salida", "COPIAR INFORMACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If (respuesta = Windows.Forms.DialogResult.Yes) Then
                            Guia_hilo.Label23.Text = "13"
                            Guia_hilo.Label25.Text = "  ALMACEN AVIOS"
                            Guia_hilo.TextBox19.Text = Label1.Text
                            Guia_hilo.Label29.Text = Label2.Text
                            Guia_hilo.TextBox1.Text = "T005"
                            Guia_hilo.Label30.Text = Label3.Text
                            Guia_hilo.Label35.Text = Trim(DataGridView1.Rows(i).Cells(3).Value)
                            Guia_hilo.Show()
                        Else
                            NsaHilo.TextBox7.Text = Label1.Text
                            NsaHilo.Label10.Text = Label2.Text
                            NsaHilo.Label13.Text = Label3.Text
                            NsaHilo.TextBox12.Text = 5
                            NsaHilo.Label21.Text = Trim(DataGridView1.Rows(i).Cells(3).Value)
                            NsaHilo.Label23.Text = Trim(DataGridView1.Rows(i).Cells(7).Value)
                            NsaHilo.TextBox8.Text = Trim(DataGridView1.Rows(i).Cells(17).Value)
                            NsaHilo.TextBox10.Text = Trim(DataGridView1.Rows(i).Cells(16).Value)
                            NsaHilo.TextBox18.Text = 5
                            NsaHilo.Show()
                        End If
                        Dim ol, ol1 As Integer
                        Dim sql102128 As String = "select COUNT(id_cabecera) from requerimiento_avios_detalle where estado1 ='1' and id_cabecera ='" + Trim(DataGridView1.Rows(i).Cells(3).Value) + "'"
                        Dim cmd102128 As New SqlCommand(sql102128, conx)
                        Rsr2128 = cmd102128.ExecuteReader()
                        If Rsr2128.Read Then
                            ol = Rsr2128(0)
                        Else
                            ol = 0
                        End If
                        Rsr2128.Close()

                        Dim sql As String = "select COUNT(id_cabecera) from requerimiento_avios_detalle where  id_cabecera ='" + Trim(DataGridView1.Rows(i).Cells(3).Value) + "'"
                        Dim sql1021299 As New SqlCommand(sql, conx)
                        Rsr21289 = sql1021299.ExecuteReader()
                        If Rsr21289.Read Then
                            ol1 = Rsr21289(0)
                        Else
                            ol1 = 0
                        End If
                        Rsr21289.Close()

                        If ol = ol1 Then
                            Dim cmd16 As New SqlCommand("UPDATE requerimiento_avios_cabecera SET estado='3' where id_requeminieto =@id_requeminieto", conx)
                            cmd16.Parameters.AddWithValue("@id_requeminieto", Trim(DataGridView1.Rows(i).Cells(3).Value))

                            cmd16.ExecuteNonQuery()

                        End If




                    End If
                End If


            Next
            MsgBox("LA INFORMACION SE ACTUALIZO CORRECTAMENTE")

        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If DataGridView1.SelectedRows.Count > 0 Then

            Dim result As DialogResult = MessageBox.Show("¿Seguro que desea Cerar la fila seleccionada?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then

                Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)


                Dim cmd161 As New SqlCommand("UPDATE requerimiento_avios_cabecera SET estado='3' where id_requeminieto =@id_requeminieto", conx)
                cmd161.Parameters.AddWithValue("@id_requeminieto", filaSeleccionada.Cells("SOLICITUD").Value)
                cmd161.ExecuteNonQuery()
                MessageBox.Show("SE CERRO EL ARMADO DE MATRIZ", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Button1.PerformClick()
            End If
        Else
            ' Si no hay una fila seleccionada, mostrar un mensaje al usuario
            MessageBox.Show("Por favor, seleccione una fila para anular.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex >= 0 Then
            Label6.Text = e.RowIndex
            Label7.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            Label8.Text = e.RowIndex
            Label9.Text = e.RowIndex
            Label10.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DataGridView1.SelectedRows.Count > 0 Then

            Dim result As DialogResult = MessageBox.Show("¿Seguro que desea Cerar la fila seleccionada?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then

                Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)


                Dim cmd161 As New SqlCommand("UPDATE requerimiento_avios_cabecera SET estado='2' where id_requeminieto =@id_requeminieto", conx)
                cmd161.Parameters.AddWithValue("@id_requeminieto", filaSeleccionada.Cells("SOLICITUD").Value)
                cmd161.ExecuteNonQuery()
                MessageBox.Show("Se Regresa el armado al estado pendiente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Button2.PerformClick()
            End If
        Else
            ' Si no hay una fila seleccionada, mostrar un mensaje al usuario
            MessageBox.Show("Por favor, seleccione una fila para anular.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If Label14.Text = "1" Then

        Else
            If RadioButton1.Checked = True Then
                buscar_op_confeccion()
            Else
                If RadioButton2.Checked = True Then
                    buscar_op_acabados()
                End If
            End If


        End If

    End Sub

    Private Sub modulo_solicitud_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button6.Visible = False
    End Sub
    Dim dg1, dg2 As New DataTable
    Sub confeccion_pendiente()
        Dim cmd As New SqlDataAdapter("select id_requeminieto AS SOLICITUD, cliente AS CLIENTE,producto AS PRODUCTO,op AS OP,po AS PO,cant_corta AS CORTADO,empresa AS EMPRESA,periodo AS PERIODO,corte  AS CORTE,'' AS 'F.ATENDIDO', '' AS 'F.DESPACHO',
            '' AS ATENDIDO, '' AS ID, '' AS TALLER , '' AS RUC, '' AS 'F.ASIGNACION'  from requerimiento_avios_cabecera where estado IN (0,1,2)  and area=1", conx)
        cmd.Fill(dg1)
        If dg1.Rows.Count <> 0 Then
            DataGridView1.DataSource = dg1
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(17).Visible = False
            DataGridView1.Columns(14).Visible = False
            DataGridView1.Columns(15).Visible = False
            DataGridView1.Columns(16).Width = 260
            DataGridView1.Columns(11).Width = 65
        Else
            DataGridView1.DataSource = ""
        End If
        Dim i51 As Integer
        i51 = DataGridView1.Rows.Count
        Dim area As String
        area = ""
        If RadioButton1.Checked = True Then
            area = "01"
        Else
            If RadioButton2.Checked = True Then
                area = "02"
            End If

        End If
        For i = 0 To i51 - 1
            If DataGridView1.Rows(i).Cells(9).Value = 2 Then
                DataGridView1.Rows(i).Cells(1).Value = True
                DataGridView1.Rows(i).Cells(1).ReadOnly = True
            Else
                DataGridView1.Rows(i).Cells(1).Value = False
            End If
            Dim sql10215 As String = "select ruc,proveedor,Fecha from asignacion_cos_aca where area ='" + area + "' and op='" + Trim(DataGridView1.Rows(i).Cells(6).Value) + "' and corte ='" + Trim(DataGridView1.Rows(i).Cells(11).Value) + "'"

            Dim cmd10215 As New SqlCommand(sql10215, conx)
            Rsr2105 = cmd10215.ExecuteReader()

            If Rsr2105.Read() Then

                DataGridView1.Rows(i).Cells(17).Value = Rsr2105(0)
                DataGridView1.Rows(i).Cells(16).Value = Rsr2105(1)
                DataGridView1.Rows(i).Cells(18).Value = Rsr2105(2)
                Rsr2105.Close()
            Else
                Rsr2105.Close()
                Dim sql10214 As String = "SELECT distinct fich_3,c.nomb_10 FROM custom_vianny.dbo.Qap3000 q 
                    inner join custom_vianny.dbo.Qag3000 g on q.Empr_3a = g.Empr_3 and q.Ano_3a=g.Ano_3 and q.talm_3a =g.talm_3 and q.ccom_3a =g.ccom_3 and q.ncom_3a = g.ncom_3  
                    left join custom_vianny.dbo.cag1000 c on q.Empr_3a = c.ccia And g.fich_3 = c.fich_10
                       where pedido_3a ='" + Trim(DataGridView1.Rows(i).Cells(6).Value) + "' and q.talm_3a ='0400'"
                Dim cmd10214 As New SqlCommand(sql10214, conx)
                Rsr210 = cmd10214.ExecuteReader()

                If Rsr210.Read() Then

                    DataGridView1.Rows(i).Cells(17).Value = Rsr210(0)
                    DataGridView1.Rows(i).Cells(16).Value = Rsr210(1)
                Else
                    DataGridView1.Rows(i).Cells(17).Value = ""
                    DataGridView1.Rows(i).Cells(16).Value = ""
                End If
                Rsr210.Close()
            End If
        Next
    End Sub
    Sub acabdos_pendiente()
        Dim cmd As New SqlDataAdapter("select id_requeminieto AS SOLICITUD, cliente AS CLIENTE,producto AS PRODUCTO,op AS OP,po AS PO,cant_corta AS CORTADO,empresa AS EMPRESA,periodo AS PERIODO,corte  AS CORTE,'' AS 'F.ATENDIDO', '' AS 'F.DESPACHO',
                '' AS ATENDIDO, '' AS ID, '' AS TALLER , '' AS RUC, '' AS 'F.ASIGNACION'  from requerimiento_avios_cabecera where estado IN (0,1,2)  and area=2", conx)
        cmd.Fill(dg2)
        If dg2.Rows.Count <> 0 Then
            DataGridView1.DataSource = dg2
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(14).Visible = False
            DataGridView1.Columns(15).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(17).Visible = False
            DataGridView1.Columns(16).Width = 260
            DataGridView1.Columns(11).Width = 65
        Else
            DataGridView1.DataSource = ""
        End If

        Dim i51 As Integer
        i51 = DataGridView1.Rows.Count

        Dim area As String
        area = ""
        If RadioButton1.Checked = True Then
            area = "01"
        Else
            If RadioButton2.Checked = True Then
                area = "02"
            End If

        End If
        For i = 0 To i51 - 1
            If DataGridView1.Rows(i).Cells(9).Value = 2 Then
                DataGridView1.Rows(i).Cells(1).Value = True
                DataGridView1.Rows(i).Cells(1).ReadOnly = True
            Else
                DataGridView1.Rows(i).Cells(1).Value = False
            End If
            Dim sql10215 As String = "select ruc,proveedor,Fecha  from asignacion_cos_aca where area ='" + area + "' and op='" + Trim(DataGridView1.Rows(i).Cells(6).Value) + "' and corte ='" + Trim(DataGridView1.Rows(i).Cells(11).Value) + "'"

            Dim cmd10215 As New SqlCommand(sql10215, conx)
            Rsr2105 = cmd10215.ExecuteReader()

            If Rsr2105.Read() Then

                DataGridView1.Rows(i).Cells(17).Value = Rsr2105(0)
                DataGridView1.Rows(i).Cells(16).Value = Rsr2105(1)
                DataGridView1.Rows(i).Cells(18).Value = Rsr2105(2)
                Rsr2105.Close()
            Else
                Rsr2105.Close()
                Dim sql10214 As String = "SELECT distinct fich_3,c.nomb_10 FROM custom_vianny.dbo.Qap3000 q 
	inner join custom_vianny.dbo.Qag3000 g on q.Empr_3a = g.Empr_3 and q.Ano_3a=g.Ano_3 and q.talm_3a =g.talm_3 and q.ccom_3a =g.ccom_3 and q.ncom_3a = g.ncom_3  
	left join custom_vianny.dbo.cag1000 c on q.Empr_3a = c.ccia And g.fich_3 = c.fich_10
    where pedido_3a ='" + Trim(DataGridView1.Rows(i).Cells(6).Value) + "' and q.talm_3a ='0500'"
                Dim cmd10214 As New SqlCommand(sql10214, conx)
                Rsr210 = cmd10214.ExecuteReader()

                If Rsr210.Read() Then

                    DataGridView1.Rows(i).Cells(17).Value = Rsr210(0)
                    DataGridView1.Rows(i).Cells(16).Value = Rsr210(1)
                Else
                    DataGridView1.Rows(i).Cells(17).Value = ""
                    DataGridView1.Rows(i).Cells(16).Value = ""
                End If
                Rsr210.Close()
            End If






        Next
    End Sub
    Friend Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Button6.Visible = False
        Button3.Enabled = True
        dg1.Clear()
        dg2.Clear()
        DataGridView1.DataSource = Nothing
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        DataGridView1.Columns(0).Visible = True
        If RadioButton1.Checked = True Then
            confeccion_pendiente()
        Else
            If RadioButton2.Checked = True Then
                acabdos_pendiente()
            End If
        End If

    End Sub



    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "SOLICITUD" Then
            rPT_SOLICITUD.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            rPT_SOLICITUD.TextBox2.Text = Label3.Text
            rPT_SOLICITUD.TextBox3.Text = Label2.Text

            rPT_SOLICITUD.Show()

        End If

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CLIENTE" Then
            detalle_requerimiento.Label1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
            detalle_requerimiento.Label2.Text = e.RowIndex
            detalle_requerimiento.Label3.Text = "2"
            detalle_requerimiento.Show()
        End If

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "ATENDIDO" Then

            AVIOS.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            AVIOS.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(12).Value
            AVIOS.TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(11).Value
            AVIOS.TextBox4.Text = "0"
            AVIOS.Show()
        End If


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        If Label7.Text = "V" Then
            MsgBox("NO HA SELECCIONADO NINGUN ITEMS")
        Else
            Detalle_ms.Label1.Text = Label7.Text
            Detalle_ms.Label4.Text = Label10.Text
            Detalle_ms.Label5.Text = Label3.Text
            Detalle_ms.Label6.Text = Label2.Text
            Detalle_ms.ShowDialog()
        End If

    End Sub

    Private Sub modulo_solicitud_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'abrir()
        'Dim p As Integer
        'p = DataGridView1.Rows.Count
        'For i = 0 To p - 1
        '    If DataGridView1.Rows(i).Cells(2).Value = True Then
        '        MsgBox(Trim(DataGridView1.Rows(i).Cells(3).Value))
        '        MsgBox(Trim(DataGridView1.Rows(i).Cells(16).Value))
        '        Dim cmd157 As New SqlCommand("update requerimiento_avios_detalle set estado1 ='0' where items IN (@id_requeminieto_detalle) AND id_cabecera =@idcabecera", conx)
        '        cmd157.Parameters.AddWithValue("@idcabecera", Trim(DataGridView1.Rows(i).Cells(3).Value))
        '        cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(i).Cells(16).Value))
        '        cmd157.ExecuteNonQuery()
        '    End If
        'Next
    End Sub

    Private Sub modulo_solicitud_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        abrir()
        Dim p As Integer
        p = DataGridView1.Rows.Count
        'For i = 0 To p - 1
        '    If DataGridView1.Rows(i).Cells(2).Value = True Then
        '        If Trim(DataGridView1.Rows(i).Cells(15).Value) <> "" Then
        '            Dim cmd1573 As New SqlCommand("update requerimiento_avios_detalle set estado1 =0 where items IN (" + DataGridView1.Rows(i).Cells(16).Value + ") AND id_cabecera =@idcabecera", conx)
        '            cmd1573.Parameters.AddWithValue("@idcabecera", Trim(DataGridView1.Rows(i).Cells(3).Value))
        '            cmd1573.ExecuteNonQuery()
        '        End If

        '    End If
        'Next

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.RowIndex >= 0 Then
            Label6.Text = e.RowIndex
            Label7.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            Label8.Text = e.RowIndex
            Label9.Text = e.RowIndex
        End If

    End Sub
End Class