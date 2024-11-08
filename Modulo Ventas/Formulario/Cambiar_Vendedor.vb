Imports System.Data.SqlClient
Public Class Cambiar_Vendedor
    Public conx As SqlConnection
    Dim ty, ty2, ty3 As New DataTable
    Public conn As SqlDataAdapter
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

    Private Sub Cambiar_Vendedor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(90)
        ty2.Clear()
        abrir()
        llenar_combo_box2()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(90)
    End Sub
    Dim dt10 As New DataTable
    Dim Rsr2 As SqlDataReader
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hj As New vcliente
        Dim jh As New fcliente

        abrir()
        Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox1.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            hj.gVendedor = Rsr2(0)

        End If

        Rsr2.Close()

        'Select Case ComboBox1.Text
        '    Case "GBEDON" : hj.gVendedor = "0010"
        '    Case "VINCIO" : hj.gVendedor = "0022"
        '    Case "DBRAVO" : hj.gVendedor = "0023"
        '    Case "JSALINAS" : hj.gVendedor = "0025"
        '    Case "GCUEVA" : hj.gVendedor = "0029"
        '    Case "AMENDO" : hj.gVendedor = "0026"
        '    Case "VGRAUS" : hj.gVendedor = "0007"
        '    Case "VPIZARRO" : hj.gVendedor = "0027"
        '    Case "GBALVIN" : hj.gVendedor = "0028"
        '    Case "VSILVERIO" : hj.gVendedor = "0005"
        '    Case "RMEDINA" : hj.gVendedor = "0030"
        '    Case "WSALINAS" : hj.gVendedor = "0034"
        'End Select
        hj.gruc = Trim(TextBox1.Text)
        hj.gfcom = DateTimePicker1.Value
        hj.gfcom_fin = DateTimePicker2.Value
        jh.actualizar_cliente(hj)

        'Dim cmd5 As New SqlDataAdapter("select COD_CLI as RUC, r_social AS CLIENTE,fcom AS FE_INICIAL, fcom_fin AS FE_FINAL, (select ISNULL( SUM(vvta1_3),0) from custom_vianny.dbo.fag0300 where fich_3 =COD_CLI  and  (fcom_3 >=fcom and fcom_3 <=fcom_fin) and flag_3 =1) as VENTA,'' AS ESTADO from cliente where Vendedor ='" + ComboBox1.SelectedValue.ToString + "'  order by VENTA desc", conx)

        'cmd5.Fill(dt10)
        'rOTACION_CLIENTE.DataGridView1.DataSource = dt10
        'rOTACION_CLIENTE.DataGridView1.Columns(3).Width = 300

        'Dim p As Integer

        'p = rOTACION_CLIENTE.DataGridView1.Rows.Count

        'For i = 0 To p - 1
        '    If rOTACION_CLIENTE.DataGridView1.Rows(i).Cells(6).Value <= 0 Then
        '        rOTACION_CLIENTE.DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
        '        rOTACION_CLIENTE.DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
        '        rOTACION_CLIENTE.DataGridView1.Rows(i).Cells(7).Value = "SIN VENTAS"
        '    Else
        '        rOTACION_CLIENTE.DataGridView1.Rows(i).Cells(7).Value = "VENTAS"
        '    End If
        'Next

        MsgBox("SE ACTUALIZO EL VENDEDOR")
        rOTACION_CLIENTE.DataGridView1.DataSource = Nothing

        Me.Close()
    End Sub
End Class