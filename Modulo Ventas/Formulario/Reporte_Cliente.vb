Imports System.Data.SqlClient
Public Class Reporte_Cliente
    Dim DT As New DataTable
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Private Sub Reporte_Cliente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box3()

        If Label4.Text = "ADMINISTRADOR" Then
            ComboBox1.Enabled = True
            'ComboBox1.SelectedIndex = 0
            abrir()
            llenar_combo_box2()

        End If

    End Sub
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2'", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2' and alias_ven ='" + MDIParent1.Label3.Text + "' ", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt10.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "RUC" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim p As Integer

                p = DataGridView1.Rows.Count
                For i = 0 To p - 1
                    DataGridView1.Columns(1).Width = 200
                    DataGridView1.Columns(2).Width = 200
                    DataGridView1.Columns(6).Width = 67
                    DataGridView1.Columns(7).Width = 67
                    If DataGridView1.Rows(i).Cells(5).Value <= 0 Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkKhaki
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Rows(i).Cells(6).Value = "NO"
                    Else
                        DataGridView1.Rows(i).Cells(6).Value = "SI"
                    End If
                    If Convert.ToDateTime(DataGridView1.Rows(i).Cells(4).Value) <= DateTimePicker1.Value.Date Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Rows(i).Cells(7).Value = "VENCIDO"
                    Else
                        DataGridView1.Rows(i).Cells(7).Value = "ACTIVO"
                    End If
                Next

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt10.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim p As Integer

                p = DataGridView1.Rows.Count
                For i = 0 To p - 1
                    DataGridView1.Columns(1).Width = 200
                    DataGridView1.Columns(2).Width = 200
                    DataGridView1.Columns(6).Width = 67
                    DataGridView1.Columns(7).Width = 67
                    If DataGridView1.Rows(i).Cells(5).Value <= 0 Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkKhaki
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Rows(i).Cells(6).Value = "NO"
                    Else
                        DataGridView1.Rows(i).Cells(6).Value = "SI"
                    End If
                    If Convert.ToDateTime(DataGridView1.Rows(i).Cells(4).Value) <= DateTimePicker1.Value.Date Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Rows(i).Cells(7).Value = "VENCIDO"
                    Else
                        DataGridView1.Rows(i).Cells(7).Value = "ACTIVO"
                    End If
                Next

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim Rsr2 As SqlDataReader
    Dim dt10 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim JK As New fcliente
        Dim JJ As New vcliente
        dt10.Clear()
        abrir()
        Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox1.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            JJ.gVendedor = Rsr2(0)

        End If

        Rsr2.Close()



        'Select Case ComboBox1.Text
        '    Case "GBEDON" : JJ.gVendedor = "0010"
        '    Case "MGRAUS" : JJ.gVendedor = "0037"
        '    Case "DBRAVO" : JJ.gVendedor = "0023"
        '    Case "JSALINAS" : JJ.gVendedor = "0025"
        '    Case "GCUEVA" : JJ.gVendedor = "0029"
        '    Case "AMENDO" : JJ.gVendedor = "0026"
        '    Case "VGRAUS" : JJ.gVendedor = "0007"
        '    Case "VPIZARRO" : JJ.gVendedor = "0027"
        '    Case "GBALVIN" : JJ.gVendedor = "0028"
        '    Case "VSILVERIO" : JJ.gVendedor = "0005"
        '    Case "WSALINAS" : JJ.gVendedor = "0034"
        'End Select
        Dim cmd5 As New SqlDataAdapter("select COD_CLI as RUC, r_social AS CLIENTE,D_fiscal AS DIRECCION,fcom AS FE_INICIAL, fcom_fin AS FE_FINAL, (select ISNULL( SUM(vvta1_3),0) from custom_vianny.dbo.fag0300 where fich_3 =COD_CLI  and  (fcom_3 >=fcom and fcom_3 <=fcom_fin) and flag_3 =1) as VENTA,'' AS VENTAS,'' AS ESTADO from cliente where Vendedor ='" + ComboBox1.SelectedValue.ToString + "'  order by VENTA desc", conx)
        cmd5.Fill(dt10)
        DataGridView1.DataSource = dt10
        Dim p As Integer

        p = DataGridView1.Rows.Count

        For i = 0 To p - 1
            DataGridView1.Columns(1).Width = 200
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(6).Width = 67
            DataGridView1.Columns(7).Width = 67
            If DataGridView1.Rows(i).Cells(5).Value <= 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Rows(i).Cells(6).Value = "NO"
            Else
                DataGridView1.Rows(i).Cells(6).Value = "SI"
            End If
            If Convert.ToDateTime(DataGridView1.Rows(i).Cells(4).Value) <= DateTimePicker1.Value.Date Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Rows(i).Cells(7).Value = "VENCIDO"
            Else
                DataGridView1.Rows(i).Cells(7).Value = "ACTIVO"
            End If
        Next

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim K As New Exportar
        K.llenarExcel(DataGridView1)
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class