Imports System.Data.SqlClient
Public Class rOTACION_CLIENTE
    Dim DT, DT2 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim ty, ty2, ty3 As New DataTable
    Public conn As SqlDataAdapter
    Dim dt10 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dt10.Clear()
        'Dim cmd5 As New SqlDataAdapter("select COD_CLI as RUC, r_social AS CLIENTE,fcom AS FE_INICIAL, fcom_fin AS FE_FINAL, (select ISNULL( SUM(vvta1_3),0) from custom_vianny.dbo.fag0300 where fich_3 =COD_CLI  and  (fcom_3 >=fcom and fcom_3 <=fcom_fin) and flag_3 =1) as VENTA,'' AS VENTAS,'' AS ESTADO from cliente where Vendedor ='" + ComboBox1.SelectedValue.ToString + "'  order by VENTA desc", conx)


        Dim cmd5 As New SqlDataAdapter("EXEC Rotacion_Clientes '" + ComboBox1.SelectedValue.ToString + "'", conx)

        cmd5.Fill(dt10)
        DataGridView1.DataSource = dt10
        DataGridView1.Columns(3).Width = 300

        Dim p As Integer

        p = DataGridView1.Rows.Count

        For i = 0 To p - 1
            If DataGridView1.Rows(i).Cells(6).Value <= 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Rows(i).Cells(7).Value = "NO"
            Else
                DataGridView1.Rows(i).Cells(7).Value = "SI"
            End If
            If Convert.ToDateTime(DataGridView1.Rows(i).Cells(5).Value) <= DateTimePicker1.Value.Date Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Rows(i).Cells(8).Value = "VENCIDO"
            Else
                DataGridView1.Rows(i).Cells(8).Value = "ACTIVO"
            End If
        Next
        ComboBox1.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
        ComboBox1.Enabled = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

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

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        DateTimePicker3.Value = DateTimePicker2.Value.AddDays(90)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim o As Integer

        o = DataGridView1.Rows.Count

        For i = 0 To o - 1
            If CheckBox1.Checked = True Then
                DataGridView1.Rows(i).Cells(0).Value = True
            Else
                DataGridView1.Rows(i).Cells(0).Value = False
            End If
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim p As Integer

        p = DataGridView1.Rows.Count

        For i = 0 To p - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then

                Dim cmd16 As New SqlCommand("update cliente set  fcom_fin =@fcom_fin  where COD_CLI =@COD_CLI and Vendedor =@Vendedor ", conx)

                cmd16.Parameters.AddWithValue("@fcom_fin", DateTimePicker3.Value.Date)
                cmd16.Parameters.AddWithValue("@COD_CLI", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd16.Parameters.AddWithValue("@Vendedor", ComboBox1.SelectedValue.ToString)
                cmd16.ExecuteNonQuery()


            End If
        Next
        DataGridView1.DataSource = Nothing
        ComboBox1.Enabled = True
        MsgBox("SE ACTUALIZO LA INFORMACION")

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim gh As New Exportar
        gh.llenarExcel(DataGridView1)
    End Sub

    Private Sub rOTACION_CLIENTE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker3.Value = DateTimePicker2.Value.AddDays(90)
        ty2.Clear()
        abrir()
        llenar_combo_box2()
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

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 1 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1 As Integer
            num1 = e.RowIndex.ToString


            Cambiar_Vendedor.TextBox1.Text = DataGridView1.Rows(num1).Cells(2).Value
            Cambiar_Vendedor.TextBox2.Text = DataGridView1.Rows(num1).Cells(3).Value
            Cambiar_Vendedor.TextBox3.Text = ComboBox1.Text
            Cambiar_Vendedor.ShowDialog()


        End If
    End Sub
End Class